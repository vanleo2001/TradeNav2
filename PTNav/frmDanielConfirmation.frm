VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmDanielConfirmation 
   Caption         =   "Form1"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrError 
      Left            =   4860
      Top             =   2640
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   3735
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
      Caption         =   "frmDanielConfirmation.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmDanielConfirmation.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDanielConfirmation.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdPark 
         Height          =   495
         Left            =   1260
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
         Caption         =   "frmDanielConfirmation.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmDanielConfirmation.frx":0092
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmDanielConfirmation.frx":00B2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   2520
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
         Caption         =   "frmDanielConfirmation.frx":00CE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmDanielConfirmation.frx":00FC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmDanielConfirmation.frx":011C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSubmit 
         Default         =   -1  'True
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
         Caption         =   "frmDanielConfirmation.frx":0138
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmDanielConfirmation.frx":0166
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmDanielConfirmation.frx":0186
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgGroups 
      Height          =   1575
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   2895
      _cx             =   5106
      _cy             =   2778
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
   Begin HexUniControls.ctlUniLabelXP lblAction 
      Height          =   255
      Left            =   180
      Top             =   180
      Width           =   4875
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
      Caption         =   "frmDanielConfirmation.frx":01A2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmDanielConfirmation.frx":0244
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDanielConfirmation.frx":0264
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraLotSize 
      Height          =   375
      Left            =   3180
      TabIndex        =   5
      Top             =   2160
      Width           =   2355
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
      Caption         =   "frmDanielConfirmation.frx":0280
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmDanielConfirmation.frx":02A0
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDanielConfirmation.frx":02C0
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniLabelXP lblLotSize 
         Height          =   195
         Left            =   0
         Top             =   83
         Width           =   1335
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
         Caption         =   "frmDanielConfirmation.frx":02DC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDanielConfirmation.frx":0324
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDanielConfirmation.frx":0344
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin gdOCX.gdScrollBar sbLotSize 
         Height          =   360
         Left            =   2160
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   635
      End
      Begin HexUniControls.ctlUniTextBoxXP txtLotSize 
         Height          =   315
         Left            =   1380
         TabIndex        =   6
         Top             =   23
         Width           =   780
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmDanielConfirmation.frx":0360
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
         Alignment       =   2
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmDanielConfirmation.frx":038A
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDanielConfirmation.frx":03AA
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraAccounts 
      Height          =   375
      Left            =   180
      TabIndex        =   8
      Top             =   2160
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
      Caption         =   "frmDanielConfirmation.frx":03C6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmDanielConfirmation.frx":03E6
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDanielConfirmation.frx":0406
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniLabelXP lblAccount 
         Height          =   195
         Left            =   0
         Top             =   83
         Width           =   735
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
         Caption         =   "frmDanielConfirmation.frx":0422
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDanielConfirmation.frx":0454
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDanielConfirmation.frx":0474
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboAccounts 
         Height          =   315
         Left            =   780
         TabIndex        =   9
         Top             =   23
         Width           =   2115
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   -2147483630
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
         Tip             =   "frmDanielConfirmation.frx":0490
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmDanielConfirmation.frx":04B0
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmDanielConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmDanielConfirmation.frm
'' Description: Form that confirms the information for the Daniel Code trades
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/05/2010   DAJ         Allow user to choose between groups
'' 08/09/2010   DAJ         Added background color change if live broker
'' 08/11/2010   DAJ         No trade on live, fix future symbols, check broker allow
'' 10/27/2010   DAJ         Implemented default group for type, Don't some symbols in grid
'' 12/08/2010   DAJ         Allow live trading with enablement code
'' 12/09/2010   DAJ         Added DoEvents to submit loop to clear up Trade Nav a bit
'' 01/26/2011   DAJ         Make sure to check correct Forex symbol when on PFG account
'' 08/21/2012   DAJ         Rename 'GmajPro' to 'DC Genie Pro' and add icons
'' 08/31/2012   DAJ         Load different DC Groups for GmajPro
'' 10/03/2012   DAJ         Lot size for forex symbols in TradeSense order groups
'' 02/01/2013   DAJ         Don't allow OK if the TSOG/Account/Symbol already exists
'' 02/05/2013   DAJ         Fix for bug where lot size always ends up being 1
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
'' 03/12/2013   DAJ         Added logging, fixed some bugs related to quantity stuff
'' 03/14/2013   DAJ         Changes related to moving Genesis Forex over to literal quantities
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kNoTradeSenseError As String = "No TradeSense Order Groups Found"
Private Const kNoTradeSymbolError As String = "You cannot trade this symbol in the selected account"
Private Const kAlreadyExistsError As String = "This TradeSense Order Group/Symbol/Account already exists"

Private Enum eGDCols
    eGDCol_On = 0
    eGDCol_Symbol
    eGDCol_Group
    eGDCol_Entry
    eGDCol_StopLoss
    eGDCol_Profit
    eGdCol_Quantity
    eGDCol_Margin
    eGDCol_Error
    eGDCol_Direction
    eGDCol_Type  ' C|R|3
    eGDCol_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user OK the dialog?
    bSubmit As Boolean                  ' Submit the groups?
    bHasForex As Boolean                ' Are there forex symbols in the grid?
    bShowLotSize As Boolean             ' Show the lot size controls?
    
    LotSize As cPriceEditor             ' Editor for the lot size
    
    DcGroups As cDanielCodeGroups       ' Collection of Daniel Code group information
End Type
Private m As mPrivate

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

Private Property Get AccountID()
    If cboAccounts.ListIndex >= 0 Then
        AccountID = cboAccounts.ItemData(cboAccounts.ListIndex)
    Else
        AccountID = -1&
    End If
End Property

Private Property Get LogPath() As String
    LogPath = AddSlash(App.Path) & "DanielCode"
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Load controls and show the form
'' Inputs:      Orders, GmajPro Version?
'' Returns:     True if Yes, False if No
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal astrOrders As cGdArray, ByVal bGmajPro As Boolean) As Boolean
On Error GoTo ErrSection:

    DumpDebug "Confirmation Form Loading: GmajPro = " & Str(bGmajPro)
    If bGmajPro Then
        Caption = "DC Genie Pro Trade Confirmation"
        Icon = Picture16(ToolbarIcon("kGmajProW"))
    Else
        Caption = "DC Genie Trade Confirmation"
        Icon = Picture16(ToolbarIcon("kDanCodeWeb"))
    End If

    Set m.DcGroups = New cDanielCodeGroups
    m.DcGroups.Load bGmajPro
    
    InitGrid
    LoadGrid astrOrders
    
    PopulateAccountsCbo cboAccounts, -1&
    
    SetupForm
    
    If fgGroups.Rows > fgGroups.FixedRows Then
        ShowForm Me, eForm_Modal, frmMain, , ALT_GRID_ROW_COLOR
        PerformAction
    Else
        DumpDebug vbTab & "No signals"
        m.bOK = False
    End If

    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmDanielConfirmation.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAccounts_Click
'' Description: Perform actions based on the selected account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAccounts_Click()
On Error GoTo ErrSection:

    If Visible Then
        DumpDebug "User changed: Account = '" & cboAccounts.Text & "' (" & Str(AccountID) & ")"
        SetupForm
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.cboAccounts_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Cancel the dialog
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
    RaiseError "frmDanielConfirmation.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPark_Click
'' Description: Park the order groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPark_Click()
On Error GoTo ErrSection:

    If (LiveAccountSelected = True) And (HasModule("TSOGLIVE") = False) Then
        InfBox "You are not authorized to submit or park these orders for a live account", "!", , "Warning"
    Else
        m.bSubmit = False
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.cmdPark_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSubmit_Click
'' Description: Submit the order groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSubmit_Click()
On Error GoTo ErrSection:

    If (LiveAccountSelected = True) And (HasModule("TSOGLIVE") = False) Then
        InfBox "You are not authorized to submit or park these orders for a live account", "!", , "Warning"
    Else
        m.bSubmit = True
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.cmdSubmit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_AfterEdit
'' Description: After user changes group, resize grid
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If Col = GDCol(eGDCol_On) Then
        DumpDebug "User changed: On(" & Str(Row) & ") = " & Str(CheckedCell(fgGroups, Row, GDCol(eGDCol_On)))
    ElseIf Col = GDCol(eGDCol_Group) Then
        CheckSymbols Row
        fgGroups.AutoSize 0, fgGroups.Cols - 1, False, 75
    
        DumpDebug "User changed: Group(" & Str(Row) & ") = " & fgGroups.TextMatrix(Row, Col)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.fgGroups_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_BeforeEdit
'' Description: Only allow the user to edit the On column
'' Inputs:      Row, Column, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim strError As String              ' Error string for the line
    Dim astrGroups As cGdArray          ' Array of groups
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrComboList As cGdArray       ' Array of group names

    If Col = GDCol(eGDCol_On) Then
        strError = fgGroups.TextMatrix(Row, GDCol(eGDCol_Error))
        If Len(strError) > 0 Then
            DumpDebug "Error turning on row " & Str(Row)
            Cancel = True
            tmrError.Tag = strError
            tmrError.Enabled = True
        End If
    ElseIf Col = GDCol(eGDCol_Group) Then
        If TypeOf fgGroups.RowData(Row) Is cGdArray Then
            Set astrGroups = fgGroups.RowData(Row)
            If astrGroups.Size <= 1 Then
                Cancel = True
            Else
                Set astrComboList = New cGdArray
                astrComboList.Create eGDARRAY_Strings
                
                For lIndex = 0 To astrGroups.Size - 1
                    astrComboList.Add Parse(astrGroups(lIndex), vbTab, 1)
                Next lIndex
                
                fgGroups.ComboList = astrComboList.JoinFields("|")
            End If
        Else
            Cancel = True
        End If
    Else
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.fgGroups_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form gets activated, move the focus to the submit button
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    MoveFocus cmdSubmit

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.Form_Activate"
    
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

    Caption = "Daniel Code Trade Confirmation"
    Icon = Picture16(ToolbarIcon("kDanCodeWeb"))
    
    g.Styler.StyleForm Me
    
    PlaceForm Me
    
    tmrError.Interval = 100
    tmrError.Enabled = False
    
    If DirExist(LogPath) = False Then
        MkDir LogPath
    End If
    KillFile AddSlash(LogPath) & "*.LOG /o=-30"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicked on the X, allow ShowMe to unload the form
'' Inputs:      Cancel Unload?, Mode of the Unload
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
    RaiseError "frmDanielConfirmation.Form_QueryUnload"
    
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

    Dim lMinScaleWidth As Long          ' Minimum scale width
    Dim lMinScaleHeight As Long         ' Minimum scale height
    Dim lSpace As Long                  ' Space between controls
    Dim lGridTop As Long                ' Top of the grid
    Dim lLeft As Long                   ' Left of the accounts control
    
    lSpace = 60

    lMinScaleWidth = fraAccounts.Width + fraLotSize.Width + (lSpace * 3)
    lMinScaleHeight = lblAction.Height + fraAccounts.Height + (fraButtons.Height * 3) + (lSpace * 3)

    If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) = False Then
        With lblAction
            .Move lSpace, lSpace
            lGridTop = .Top + .Height
        End With
        
        With fraButtons
            .Move (ScaleWidth / 2) - (.Width / 2), ScaleHeight - .Height - lSpace
        End With
        
        If m.bShowLotSize Then
            lLeft = (ScaleWidth / 2) - ((fraAccounts.Width + fraLotSize.Width + lSpace) / 2)
        Else
            lLeft = (ScaleWidth / 2) - (fraAccounts.Width / 2)
        End If
        
        With fraAccounts
            .Move lLeft, fraButtons.Top - .Height - lSpace
        End With
        
        With fraLotSize
            .Move lLeft + fraAccounts.Width + lSpace, fraAccounts.Top
        End With
        
        With fgGroups
            .Move lSpace, lGridTop, ScaleWidth - (lSpace * 2), fraAccounts.Top - lGridTop
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up and save settings when form is unloaded
'' Inputs:      Cancel Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    tmrError.Enabled = False
    
    SaveFormPlacement Me
    SetIniFileProperty "LastLotSize", m.LotSize.Price, "DanielConfirmation", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrError_Timer
'' Description: Display the error message
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrError_Timer()
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message to display to the user

    tmrError.Enabled = False
    strMessage = "You cannot activate this group because|" & tmrError.Tag & "|"
    InfBox strMessage, "!", , "Error"
    DumpDebug "Error shown to user: " & Replace(strMessage, "|", " '")
    tmrError.Tag = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.tmrError_Timer"
    
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
        
        SetupGrid fgGroups, eGridMode_Grid
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_On)) = "On"
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_Group)) = "Group"
        .TextMatrix(0, GDCol(eGDCol_Entry)) = "Entry Price"
        .TextMatrix(0, GDCol(eGDCol_StopLoss)) = "Stop Loss"
        .TextMatrix(0, GDCol(eGDCol_Profit)) = "Profit Target"
        .TextMatrix(0, GDCol(eGdCol_Quantity)) = "Quantity"
        .TextMatrix(0, GDCol(eGDCol_Margin)) = "Margin"
        .TextMatrix(0, GDCol(eGDCol_Error)) = "Error"
        .TextMatrix(0, GDCol(eGDCol_Direction)) = "Dir"
        .TextMatrix(0, GDCol(eGDCol_Type)) = "Type"
        
        .ColDataType(GDCol(eGDCol_On)) = flexDTBoolean
        .ColHidden(GDCol(eGDCol_Error)) = True
        .ColHidden(GDCol(eGDCol_Direction)) = True
        .ColHidden(GDCol(eGDCol_Type)) = True
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid
'' Inputs:      Orders
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid(ByVal astrOrders As cGdArray)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrOrder As cGdArray           ' Information for an order split out into an array
    Dim astrGroups As cGdArray          ' Groups for the direction/symbol
    Dim lDefault As Long                ' Default index for given type
    Dim strSymbol As String             ' Fixed symbol
    
    Set astrOrder = New cGdArray
    
    DumpDebug "LoadingGrid()"
    With fgGroups
        .Redraw = flexRDNone
        
        m.bHasForex = False
        For lIndex = 0 To astrOrders.Size - 1
            If Len(Trim(astrOrders(lIndex))) > 0 Then
                DumpDebug vbTab & "Order (" & Str(lIndex) & "): " & astrOrders(lIndex)
                
                astrOrder.Clear
                astrOrder.SplitFields astrOrders(lIndex), vbTab
                
                ' 10/27/2010 DAJ: Only add it to the grid if the symbol is in the symbol pool and
                ' if it is not SP3, ZD, or ND3...
                If g.SymbolPool.PoolRecForSymbol(FixSymbol(astrOrder(0))) >= 0 Then
                    If (Parse(astrOrder(0), "-", 1) <> "SP3") And (Parse(astrOrder(0), "-", 1) <> "ZD") And (Parse(astrOrder(0), "-", 1) <> "ND3") Then
                        .Rows = .Rows + 1
                        
                        Set astrGroups = m.DcGroups.GetGroups(astrOrder(1), FixSymbol(astrOrder(0)))
                        If astrGroups.Size > 0 Then
                            .RowData(.Rows - 1) = astrGroups
                            
                            lDefault = m.DcGroups.GetDefault(astrOrder(1), FixSymbol(astrOrder(0)), astrOrder(7))
                            
                            CheckedCell(fgGroups, .Rows - 1, GDCol(eGDCol_On)) = True
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Group)) = Parse(astrGroups(lDefault), vbTab, 1)
                            
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Error)) = ""
                            .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = .Cell(flexcpForeColor, 0, 0)
                        Else
                            CheckedCell(fgGroups, .Rows - 1, GDCol(eGDCol_On)) = False
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Group)) = ""
                            
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Error)) = kNoTradeSenseError
                            .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                        End If
                        
                        strSymbol = FixSymbol(astrOrder(0))
                        If IsForex(strSymbol) Then
                            m.bHasForex = True
                        End If
                    
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = strSymbol
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Entry)) = astrOrder(2)
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_StopLoss)) = astrOrder(3)
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Profit)) = astrOrder(4)
                        .TextMatrix(.Rows - 1, GDCol(eGdCol_Quantity)) = astrOrder(5)
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Margin)) = astrOrder(6)
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Direction)) = astrOrder(1)
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Type)) = astrOrder(7)
                        
                        DumpDebug vbTab & "To Grid (" & Str(.Rows - 1) & "): " & RowToString(.Rows - 1, vbTab)
                    End If
                End If
            End If
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.LoadGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PerformAction
'' Description: Either Submit or Park the groups
'' Inputs:      Orders
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PerformAction()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    Dim strGroupInfo As String          ' Group information for the given group name
    Dim tsoGrp As cTradeSenseOrderGroup ' TradeSense order group object
    Dim tsGroup As cActiveTsOrderGroup  ' Active TradeSense order group object
    Dim astrInfo As cGdArray            ' Group information
    
    DumpDebug "PerformAction(OK = '" & Str(m.bOK) & "', Submit = '" & Str(m.bSubmit) & "', Account = '" & cboAccounts.Text & "', LotSize = '" & Str(m.LotSize.Price) & "')"
    If m.bOK Then
        With fgGroups
            For lIndex = .FixedRows To .Rows - 1
                DumpDebug vbTab & "From Grid (" & Str(lIndex) & "): " & RowToString(lIndex, vbTab)
                
                If CheckedCell(fgGroups, lIndex, GDCol(eGDCol_On)) = True Then
                    Set astrInfo = New cGdArray
                    astrInfo.Create eGDARRAY_Strings
                    
                    strGroupInfo = m.DcGroups.GetGroupInfo(.TextMatrix(lIndex, GDCol(eGDCol_Direction)), .TextMatrix(lIndex, GDCol(eGDCol_Symbol)), .TextMatrix(lIndex, GDCol(eGDCol_Group)))
                    If Len(strGroupInfo) > 0 Then
                        Set tsoGrp = New cTradeSenseOrderGroup
                        If tsoGrp.LoadDanielCodeGroup(strGroupInfo, .TextMatrix(lIndex, GDCol(eGDCol_Entry)), .TextMatrix(lIndex, GDCol(eGDCol_StopLoss)), .TextMatrix(lIndex, GDCol(eGDCol_Profit))) Then
                            Set tsGroup = New cActiveTsOrderGroup
                            
                            astrInfo.Add "Symbol = '" & .TextMatrix(lIndex, GDCol(eGDCol_Symbol)) & "'"
                            tsGroup.SymbolOrSymbolID = .TextMatrix(lIndex, GDCol(eGDCol_Symbol))
                            
                            astrInfo.Add "Quantity = '" & .TextMatrix(lIndex, GDCol(eGdCol_Quantity)) & "'"
                            tsGroup.Quantity = CLng(Val(.TextMatrix(lIndex, GDCol(eGdCol_Quantity))))
                            
                            If (IsForex(tsGroup.Symbol) = True) And (m.bShowLotSize = True) Then
                                astrInfo.Add "LotSize = '" & Str(m.LotSize.Price) & "'"
                                tsGroup.LotSize = m.LotSize.Price
                            Else
                                astrInfo.Add "LotSize = '1'"
                                tsGroup.LotSize = 1&
                            End If
                            
                            astrInfo.Add "Account = '" & Str(AccountID) & "'"
                            tsGroup.AccountID = AccountID
                            
                            astrInfo.Add "Group = '" & tsoGrp.Name & "'"
                            tsGroup.tsOrderGroup = tsoGrp

                            If m.bSubmit Then
                                DumpDebug vbTab & vbTab & "Submit: " & Chr(34) & astrInfo.JoinFields(vbTab) & Chr(34)
                                g.TsoGroups.Submit tsGroup
                            Else
                                DumpDebug vbTab & vbTab & "Park: " & Chr(34) & astrInfo.JoinFields(vbTab) & Chr(34)
                                g.TsoGroups.Park tsGroup
                            End If
                        End If
                    End If
                End If
                
                DoEvents
            Next lIndex
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.PerformAction"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixSymbol
'' Description: Fix the symbol if it is a Genesis forex symbol
'' Inputs:      Symbol
'' Returns:     Fixed symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FixSymbol(ByVal strSymbol As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    strReturn = strSymbol
    If IsForex("$" & strSymbol) = True Then
        If IsAlpha(strSymbol, 5) Then
            strReturn = "$" & strSymbol
        End If
    End If
    
    If SecurityType(strReturn, True) = "F" Then
        strSymbol = ConvertFutureSymbol(strReturn, eElectronicSymbol)
        If (Len(strSymbol) > 0) Then
            strReturn = strSymbol
        End If
    End If
    
    FixSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDanielConfirmation.FixSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColorForm
'' Description: Color the background color of the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ColorForm()
On Error GoTo ErrSection:

    Dim lBackColor As Long              ' Background color
    Dim bLiveAccount As Boolean         ' Live account selected?
    Static bPrevLiveAccount As Boolean  ' Was a live account selected?
    
    bLiveAccount = LiveAccountSelected
    If bLiveAccount <> bPrevLiveAccount Then
        If bLiveAccount Then
            lBackColor = kFrameLive
        Else
            lBackColor = GetAppBackColor
        End If
        
        BackColor = lBackColor
        lblAccount.BackColor = lBackColor
        lblAction.BackColor = lBackColor
        fraAccounts.BackColor = lBackColor
        fraLotSize.BackColor = lBackColor
        lblLotSize.BackColor = lBackColor
        
        bPrevLiveAccount = bLiveAccount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.ColorForm"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LiveAccountSelected
'' Description: Is the currently selected account a live account?
'' Inputs:      None
'' Returns:     True if live account selected, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LiveAccountSelected() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If cboAccounts.ListIndex > -1 Then
        bReturn = (TypeOfAccount(AccountID) = eGDTypeOfAccount_BrokerLive)
    End If
    
    LiveAccountSelected = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDanielConfirmation.LiveAccountSelected"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckSymbols
'' Description: Check to make sure symbols are valid for the selected account
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckSymbols(Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    Dim lAccountID As Long              ' Account ID from the combo box
    Dim nBroker As eTT_AccountType      ' Broker for the account
    Dim strSymbol As String             ' Symbol to check
    Dim strTsogId As String             ' TradeSense Order Group ID
    Dim strKey As String                ' Active TradeSense Order Group Key
    Dim lFrom As Long                   ' Starting row for the for loop
    Dim lTo As Long                     ' Ending row for the for loop
    
    If cboAccounts.ListIndex > -1 Then
        lAccountID = AccountID
        nBroker = g.Broker.AccountTypeForID(lAccountID)
        
        With fgGroups
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            If lRow = -1 Then
                lFrom = .FixedRows
                lTo = .Rows - 1
            Else
                lFrom = lRow
                lTo = lRow
            End If
            
            For lIndex = lFrom To lTo
                If .TextMatrix(lIndex, GDCol(eGDCol_Error)) <> kNoTradeSenseError Then
                    strSymbol = .TextMatrix(lIndex, GDCol(eGDCol_Symbol))
                    
                    ' 01/26/2011 DAJ: If this is a PFG account and the symbol is a Forex symbol,
                    ' make sure that it has the @PFG when checking if we can trade it...
                    strSymbol = FixForexSymbol(strSymbol)
                    strTsogId = TsogIdForRow(lIndex)
                    strKey = strSymbol & vbTab & Str(lAccountID) & vbTab & strTsogId
                    
                    If g.TsoGroups.Exists(strKey) Then
                        CheckedCell(fgGroups, lIndex, GDCol(eGDCol_On)) = False
                        .TextMatrix(lIndex, GDCol(eGDCol_Error)) = kAlreadyExistsError
                        .Cell(flexcpForeColor, lIndex, 0, lIndex, .Cols - 1) = vbRed
                    ElseIf g.Broker.CanTrade(lAccountID, strSymbol) Then
                        CheckedCell(fgGroups, lIndex, GDCol(eGDCol_On)) = True
                        .TextMatrix(lIndex, GDCol(eGDCol_Error)) = ""
                        .Cell(flexcpForeColor, lIndex, 0, lIndex, .Cols - 1) = .Cell(flexcpForeColor, 0, 0)
                    Else
                        CheckedCell(fgGroups, lIndex, GDCol(eGDCol_On)) = False
                        .TextMatrix(lIndex, GDCol(eGDCol_Error)) = kNoTradeSymbolError
                        .Cell(flexcpForeColor, lIndex, 0, lIndex, .Cols - 1) = vbRed
                    End If
                End If
            Next lIndex
            
            .Redraw = nRedraw
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.CheckSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetupLotSize
'' Description: Setup the lot size controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupLotSize()
On Error GoTo ErrSection:

    Dim lLastLotSize As Long            ' Last lot size used
    
    lLastLotSize = GetIniFileProperty("LastLotSize", 10000&, "DanielConfirmation", g.strIniFile)
    
    Set m.LotSize = New cPriceEditor
    txtLotSize.Text = Str(lLastLotSize)
    m.LotSize.Init sbLotSize, txtLotSize, Nothing, lLastLotSize, 10000
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.SetupLotSize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowLotSizeControls
'' Description: Show/Hide the lot size controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowLotSizeControls()
On Error GoTo ErrSection:

    m.bShowLotSize = False
    If cboAccounts.ListIndex >= 0 Then
        m.bShowLotSize = (m.bHasForex = True) And (m.LotSize.Min > 1)
    End If
    
    fraLotSize.Visible = m.bShowLotSize
    Form_Resize

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.ShowLotSizeControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TsogIdForRow
'' Description: Determine the TradeSense Order Group ID for the given row
'' Inputs:      Row
'' Returns:     TradeSense Order Group ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TsogIdForRow(ByVal lRow As Long) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strGroupInfo As String          ' Group information for the given group name
    Dim tsoGrp As cTradeSenseOrderGroup ' TradeSense order group object

    With fgGroups
        If (lRow >= .FixedRows) And (lRow < .Rows) Then
            strGroupInfo = m.DcGroups.GetGroupInfo(.TextMatrix(lRow, GDCol(eGDCol_Direction)), .TextMatrix(lRow, GDCol(eGDCol_Symbol)), .TextMatrix(lRow, GDCol(eGDCol_Group)))
            If Len(strGroupInfo) > 0 Then
                Set tsoGrp = New cTradeSenseOrderGroup
                If tsoGrp.LoadDanielCodeGroup(strGroupInfo, .TextMatrix(lRow, GDCol(eGDCol_Entry)), .TextMatrix(lRow, GDCol(eGDCol_StopLoss)), .TextMatrix(lRow, GDCol(eGDCol_Profit))) Then
                    strReturn = tsoGrp.ID
                End If
            End If
        End If
    End With

    TsogIdForRow = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDanielConfirmation.TsogIdForRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixForexSymbol
'' Description: Fix the given forex symbol for the given account
'' Inputs:      Forex symbol
'' Returns:     Fixed forex symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FixForexSymbol(ByVal strSymbol As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim nBroker As eTT_AccountType      ' Broker
    Dim strSuffix As String             ' Suffix for the symbol
    
    strReturn = strSymbol
    If (IsForex(strSymbol) = True) And (AccountID <> -1&) Then
        nBroker = g.Broker.AccountTypeForID(AccountID)
        strSuffix = ""
        
        If g.Broker.IsPfgBroker(nBroker) Then
            strSuffix = "@PFG"
        ElseIf g.Broker.IsCurrenexBroker(nBroker) Then
            strSuffix = "@CNX"
        ElseIf g.Broker.IsIbBroker(nBroker) Then
            strSuffix = "@IB"
        End If
        
        strReturn = Parse(strSymbol, "@", 1) & strSuffix
    End If
    
    FixForexSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDanielConfirmation.FixForexSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SampleForexSymbol
'' Description: Build a sample forex symbol for the given account
'' Inputs:      None
'' Returns:     Sample forex symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SampleForexSymbol() As String
On Error GoTo ErrSection:

    SampleForexSymbol = FixForexSymbol("$EUR-USD")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDanielConfirmation.SampleForexSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitQuantityEditor
'' Description: Initialize the quantity editor according to the selected
''              account and symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitQuantityEditor()
On Error GoTo ErrSection:

    g.Broker.InitQuantityEditor m.LotSize, sbLotSize, txtLotSize, AccountID, SampleForexSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.InitQuantityEditor"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetupForm
'' Description: Setup the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupForm()
On Error GoTo ErrSection:

    ColorForm
    CheckSymbols
    If m.LotSize Is Nothing Then
        SetupLotSize
    End If
    
    ' 03/13/2013 DAJ: Make sure to call InitQuantityEditor before ShowLotSizeControls
    ' because we are going to use the min value on the editor to determine whether or
    ' not to show the lot size controls...
    InitQuantityEditor
    ShowLotSizeControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDanielConfirmation.SetupForm"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowToString
'' Description: Convert the given row to a delimited string
'' Inputs:      Row, Delimiter
'' Returns:     Delimited string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RowToString(ByVal lRow As Long, ByVal strDelimiter As String) As String
On Error GoTo ErrSection:

    Dim lCol As Long                    ' Index into a for loop
    Dim astrCols As cGdArray            ' Array of column information from the grid

    Set astrCols = New cGdArray
    astrCols.Create eGDARRAY_Strings
    
    With fgGroups
        If (lRow >= .FixedRows) And (lRow < .Rows) Then
            For lCol = 0 To .Cols - 1
                If lCol = GDCol(eGDCol_On) Then
                    astrCols.Add "'" & Str(CheckedCell(fgGroups, lRow, lCol)) & "'"
                Else
                    astrCols.Add "'" & .TextMatrix(lRow, lCol) & "'"
                End If
            Next lCol
        End If
    End With
    
    RowToString = astrCols.JoinFields(strDelimiter)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmDanielConfirmation.RowToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Send a string to the log file for the day
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpDebug(ByVal strMessage As String)
On Error Resume Next

#If 0 Then

    Dim fh As Integer                   ' File handle to open file with
    fh = FreeFile
    Open AddSlash(LogPath) & "TN" & Format(Now, "YYYYMMDD") & ".LOG" For Append As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " (" & Str(gdTickCount) & ") - " & strMessage
        Close #fh
    End If

#Else

    Static LogFile As cLogFile
    If LogFile Is Nothing Then
        Set LogFile = New cLogFile
        LogFile.OpenFile AddSlash(LogPath) & "TN*.LOG"
    End If
    LogFile.WriteText strMessage

#End If

End Sub

