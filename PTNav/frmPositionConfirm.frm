VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmPositionConfirm 
   Caption         =   "Form1"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPbo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2640
      Picture         =   "frmPositionConfirm.frx":0000
      ScaleHeight     =   210
      ScaleWidth      =   1830
      TabIndex        =   4
      Top             =   4980
      Width           =   1830
   End
   Begin VB.PictureBox picRithmic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   60
      Picture         =   "frmPositionConfirm.frx":050E
      ScaleHeight     =   345
      ScaleWidth      =   1995
      TabIndex        =   5
      Top             =   4920
      Width           =   1995
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   4320
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
      Caption         =   "frmPositionConfirm.frx":079D
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPositionConfirm.frx":07C9
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPositionConfirm.frx":07E9
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDisconnect 
         Height          =   495
         Left            =   1380
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
         Caption         =   "frmPositionConfirm.frx":0805
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPositionConfirm.frx":083B
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPositionConfirm.frx":085B
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   0
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
         Caption         =   "frmPositionConfirm.frx":0877
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmPositionConfirm.frx":089D
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmPositionConfirm.frx":08BD
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgPositions 
      Height          =   2295
      Left            =   60
      TabIndex        =   2
      Top             =   1080
      Width           =   4395
      _cx             =   7752
      _cy             =   4048
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
      Rows            =   1
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
   Begin HexUniControls.ctlUniLabelXP lblConfirm 
      Height          =   615
      Left            =   120
      Top             =   3540
      Width           =   4335
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
      Caption         =   "frmPositionConfirm.frx":08D9
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmPositionConfirm.frx":0A57
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPositionConfirm.frx":0A77
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblIntro 
      Height          =   795
      Left            =   120
      Top             =   120
      Width           =   4395
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
      Caption         =   "frmPositionConfirm.frx":0A93
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmPositionConfirm.frx":0C4D
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPositionConfirm.frx":0C6D
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmPositionConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmPositionConfirm.frm
'' Description: Allow the user to confirm their positions for live trading
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/11/2010   DAJ         Use global Order Strategies, Trading Items, DoBrokerTimer
'' 09/24/2010   DAJ         Added some artwork to be shown for Rithmic
'' 10/05/2010   DAJ         Changed the Rithmic image
'' 10/26/2010   DAJ         Changed the Rithmic image
'' 10/27/2010   DAJ         More mods to the Rithmic image
'' 01/06/2011   DAJ         Check account when sending position to auto-trade item
'' 05/11/2011   DAJ         Utilize IsLiveAccount function
'' 06/21/2011   DAJ         Separate out Simulated trading types
'' 07/15/2011   DAJ         Enhancements for allowing auto trading on continuous contracts
'' 07/19/2011   DAJ         Editing automated position now calls reassign fills form
'' 07/21/2011   DAJ         Added ShowMeForTradeItem, changed call to trade item enable
'' 09/15/2011   DAJ         Fix for activating auto trade item for synthetic symbol
'' 01/17/2012   DAJ         Allow auto trade enable on load when in mismatch
'' 01/18/2012   DAJ         Enhanced logging for automated trading
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 12/13/2012   DAJ         Have Ctrl-Click enable/disable all automated trading items
'' 12/13/2012   DAJ         Do a DoEvents every 10 times throught the ActivateItems loop
'' 01/18/2013   DAJ         Don't allow automated trading for spreads
'' 02/06/2013   DAJ         Fix for not being able to turn on automated trading item
'' 04/03/2013   DAJ         Automated Strategy Baskets
'' 05/01/2013   DAJ         Don't allow enable of auto trade item if no longer authorized
'' 05/24/2013   DAJ         Speed enhancements
'' 06/11/2013   DAJ         Clean out variables on unload of the form
'' 06/12/2013   DAJ         Symbol and quantity validation for automated trading items
'' 07/08/2014   DAJ         Don't automatically disable auto trade item if max units is zero
'' 10/22/2014   DAJ         Do a DoEvents everty time through the ActivateItems loop
'' 11/11/2014   DAJ         Mark automated trading item enabled if "previously active"
'' 11/11/2014   DAJ         Dump the status of the enabled column in the DumpGrid routine
'' 12/04/2014   DAJ         Removed old code; Don't show enabled error if this form not shown
'' 02/26/2015   DAJ         When activating auto trade items, pass the current position from BInfo not the grid
'' 01/15/2016   DAJ         Added "All Automated Trading Items" row to the grid
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?
    bModal As Boolean                   ' Is this form being shown modally?
    bShow As Boolean                    ' Is the form going to be shown?
    nBroker As eTT_AccountType          ' Broker for this instance of the form
    strAccount As String                ' Account number to filter on
    strSymbol As String                 ' Symbol to filter on
    
    bTradeItemMode As Boolean           ' Did we get called in Trade Item mode?
    TradeItem As cAutoTradeItem         ' Auto Trade Item
    astrTradeItemSymbols As cGdArray    ' Trade Item symbols
    astrTradeItemAccounts As cGdArray   ' Trade Item accounts
    alEnableAcctPosIds As cGdArray      ' Trade Item account position ID's
End Type
Private m As mPrivate

Private Enum eGDCols
    eGDCol_Account = 0
    eGDCol_Symbol
    eGDCol_Position
    eGDCol_Buys
    eGDCol_Sells
    eGDCol_Overnight
    eGDCol_Source
    eGDCol_SourceID
    eGDCol_Enable
    eGDCol_SymbolError
    eGDCol_QuantityError
    
    eGDCol_NumCols
End Enum

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Broker, Account, Symbol, Modal?, Show?
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal nBroker As eTT_AccountType, Optional ByVal strAccount As String = "", Optional ByVal strSymbol As String = "", Optional ByVal bModal As Boolean = False, Optional ByVal bShow As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bRithmicBroker As Boolean       ' Is this a Rithmic broker?
    Static bInProgress As Boolean       ' Are we already confirming a position?
    
    If bInProgress = False Then
        bInProgress = True
    
        m.bModal = bModal
        m.bShow = bShow
        m.nBroker = nBroker
        m.strAccount = strAccount
        
        If m.bTradeItemMode = False Then
            Set m.TradeItem = Nothing
            Set m.astrTradeItemSymbols = New cGdArray
            m.astrTradeItemSymbols.Create eGDARRAY_Strings
            Set m.astrTradeItemAccounts = New cGdArray
            m.astrTradeItemAccounts.Create eGDARRAY_Strings
            Set m.alEnableAcctPosIds = New cGdArray
            m.alEnableAcctPosIds.Create eGDARRAY_Longs
        End If
        
        If Len(strSymbol) > 0 Then
            m.strSymbol = ConvertToTradeSymbol(strSymbol)
        Else
            m.strSymbol = strSymbol
        End If
        
        If Not m.TradeItem Is Nothing Then
            Caption = "Position Confirmation for " & m.TradeItem.Name
        ElseIf Len(m.strAccount) > 0 And Len(m.strSymbol) > 0 Then
            Caption = "Position Confirmation for " & strSymbol & " in Account " & strAccount
        ElseIf Len(m.strAccount) > 0 Then
            Caption = "Position Confirmation for Account " & strAccount
        ElseIf Len(m.strSymbol) > 0 Then
            Caption = "Position Confirmation for " & strSymbol & " through " & g.Broker.BrokerName(nBroker)
        Else
            Caption = "Position Confirmation for " & g.Broker.BrokerName(nBroker)
        End If
        
        InitGrid
        LoadGrid
        
        If bShow Then
            If fgPositions.Rows > fgPositions.FixedRows Then
                Enable cmdDisconnect, g.Broker.IsLiveAccount(m.nBroker)
                bRithmicBroker = g.Broker.IsRithmicBroker(m.nBroker)
                picRithmic.Visible = bRithmicBroker
                picPbo.Visible = bRithmicBroker
                
                Form_Resize
                If bModal Then
                    ' Don't do an ActModal here because it can conflict on startup with the message box
                    ' asking if the user wants to recalculate criteria...
                    ShowForm Me, eForm_Modal, frmMain, , ALT_GRID_ROW_COLOR
                    SaveResults
                    ShowMe = m.bOK
                Else
                    ShowForm Me, eForm_Nonmodal, frmMain
                    ShowMe = True
                End If
            Else
                If g.Broker.IsLiveAccount(m.nBroker) Then
                    InfBox "There is no activity on " & g.Broker.BrokerName(nBroker) & "|to verify at this time.", "i", "+-OK", Caption
                End If
                m.bOK = True
                SaveResults
                ShowMe = True
                If Not bModal Then Unload Me
            End If
        Else
            m.bOK = True
            SaveResults
            ShowMe = True
            If Not bModal Then Unload Me
        End If
        bInProgress = False
    End If
        
ErrExit:
    If ((bModal = True) And (bInProgress = False)) Then
        Unload Me
    End If
    Exit Function
    
ErrSection:
    If bModal Then Unload Me
    bInProgress = False
    RaiseError "frmPositionConfirm.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeForTradeItem
'' Description: Setup and show the form for the given trading item
'' Inputs:      Trade Item
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeForTradeItem(ByVal TradeItem As cAutoTradeItem) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim ChildItem As cAutoTradeItem     ' Child auto trade item
    
    Set m.TradeItem = TradeItem
    Set m.astrTradeItemSymbols = New cGdArray
    m.astrTradeItemSymbols.Create eGDARRAY_Strings
    Set m.astrTradeItemAccounts = New cGdArray
    m.astrTradeItemAccounts.Create eGDARRAY_Strings
    Set m.alEnableAcctPosIds = New cGdArray
    m.alEnableAcctPosIds.Create eGDARRAY_Longs
    
    If m.TradeItem.ParentID = -1& Then
        For lIndex = 1 To g.TradingItems.Count
            Set ChildItem = g.TradingItems(lIndex)
            If ChildItem.ParentID = TradeItem.AutoTradeItemID Then
                AddTradeItem ChildItem
            End If
        Next lIndex
    Else
        AddTradeItem TradeItem
    End If
    
    m.bTradeItemMode = True
    ShowMeForTradeItem = ShowMe(TradeItem.Broker, , , True)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPositionConfirm.ShowMeForTradeItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDisconnect_Click
'' Description: User did not OK the dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDisconnect_Click()
On Error GoTo ErrSection:

    g.Broker.Disconnect m.nBroker, "Did not verify positions"

    If m.bModal Then
        m.bOK = False
        Hide
    Else
        Unload Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.cmdDisconnect_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: User OKed the dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    
    If m.bModal Then
        Hide
    Else
        Hide
        SaveResults
        Unload Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgPositions_AfterEdit
'' Description: Handle the user's changes
'' Inputs:      Row and Column of Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgPositions_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim lAutoTradeItemID As Long        ' Automated trading item ID
    Dim strSymbolError As String        ' Symbol error
    Dim strQtyError As String           ' Quantity error

    With fgPositions
        Select Case Col
            Case GDCol(eGDCol_Position)
                'RecalcChildren .GetNodeRow(Row, flexNTParent)
                
            Case GDCol(eGDCol_Source)
                If .TextMatrix(Row, Col) = "None" Then
                    CheckedCell(fgPositions, Row, GDCol(eGDCol_Enable)) = False
                'ElseIf VerifyAutoExits(.RowData(Row).AccountID, .RowData(Row).SymbolOrSymbolID) Then
                ElseIf CanActivateAutomatedItem(.RowData(Row).AccountID, .RowData(Row).SymbolOrSymbolID, "Auto Exit", "Position Confirm") Then
                    CheckedCell(fgPositions, Row, GDCol(eGDCol_Enable)) = True
                Else
                    CheckedCell(fgPositions, Row, GDCol(eGDCol_Enable)) = False
                    .TextMatrix(Row, Col) = "None"
                End If
                
            Case GDCol(eGDCol_Enable)
                lAutoTradeItemID = CLng(Val(.TextMatrix(Row, GDCol(eGDCol_SourceID))))
                
                If .MergeRow(Row) = True Then
                    ToggleAllAutoTradeItems CheckedCell(fgPositions, Row, Col)
                
                ElseIf lAutoTradeItemID = 0& Then
                    If CheckedCell(fgPositions, Row, Col) = True Then
                        If CanActivateAutomatedItem(.RowData(Row).AccountID, .RowData(Row).SymbolOrSymbolID, "Auto Exit", "Position Confirm") Then
                            .Col = GDCol(eGDCol_Source)
                        Else
                            CheckedCell(fgPositions, Row, Col) = False
                        End If
                    Else
                        .TextMatrix(Row, GDCol(eGDCol_Source)) = "None"
                    End If
                
                ElseIf lAutoTradeItemID > 0& Then
                    ToggleAutoTradeItem Row, CheckedCell(fgPositions, Row, Col), True
                    
                    If AllAutoTradeItemsAre(True) Then
                        CheckedCell(fgPositions, .FixedRows, GDCol(eGDCol_Enable)) = True
                    ElseIf AllAutoTradeItemsAre(False) Then
                        CheckedCell(fgPositions, .FixedRows, GDCol(eGDCol_Enable)) = False
                    End If
                    
'                    If CheckedCell(fgPositions, Row, Col) = True Then
'                        strSymbolError = .TextMatrix(Row, GDCol(eGDCol_SymbolError))
'                        strQtyError = .TextMatrix(Row, GDCol(eGDCol_QuantityError))
'
'                        If Len(strSymbolError) > 0 Then
'                            CheckedCell(fgPositions, Row, Col) = False
'                            InfBox strSymbolError, "!", , "Error"
'                        ElseIf Len(strQtyError) > 0 Then
'                            CheckedCell(fgPositions, Row, Col) = False
'                            InfBox strQtyError, "!", , "Error"
'                        ElseIf mSysNav.MaxUnitsForAutoTradeID(lAutoTradeItemID) = 0 Then
'                            CheckedCell(fgPositions, Row, Col) = False
'                        ElseIf CanActivateAutomatedItem(.RowData(Row).AccountID, .RowData(Row).SymbolOrSymbolID, "Automated Trading Item", "Position Confirm") = False Then
'                            CheckedCell(fgPositions, Row, Col) = False
'                        End If
'                    End If
                End If
                
        End Select
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.fgPositions_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgPositions_AfterRowColChange
'' Description: Turn the edit cell on after a row or column change
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgPositions_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    'If NewCol <> GDCol(eGDCol_Enable) Then
    '    fgPositions.EditCell
    'End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.fgPositions_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgPositions_BeforeEdit
'' Description: Only allow the user to edit certain columns and cells
'' Inputs:      Row and Column of Edit, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgPositions_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    fgPositions.ComboList = ""
    If Row < fgPositions.FixedRows Then
        Cancel = True
    Else
        Select Case Col
            Case GDCol(eGDCol_Account), GDCol(eGDCol_Symbol), GDCol(eGDCol_Buys), GDCol(eGDCol_Sells), GDCol(eGDCol_Overnight), GDCol(eGDCol_SourceID)
                Cancel = True
            
            Case GDCol(eGDCol_Position)
                If Val(fgPositions.TextMatrix(Row, GDCol(eGDCol_SourceID))) <= 0# Then
                    Cancel = True
                Else
                    fgPositions.ComboList = "..."
                End If
                
            Case GDCol(eGDCol_Source)
                If Val(fgPositions.TextMatrix(Row, GDCol(eGDCol_SourceID))) <> 0# Then
                    Cancel = True
                ElseIf AllowAutoExits(fgPositions.RowData(Row).AccountID, fgPositions.RowData(Row).SymbolOrSymbolID, False) Then
                    fgPositions.ComboList = AutoExitList
                Else
                    Cancel = True
                End If
                
            Case GDCol(eGDCol_Enable)
                If fgPositions.TextMatrix(Row, GDCol(eGDCol_SourceID)) = "-1" Then
                    Cancel = True
                ElseIf UCase(fgPositions.TextMatrix(Row, GDCol(eGDCol_Source))) = "NONE" Then
                    Cancel = True
                    If AllowAutoExits(fgPositions.RowData(Row).AccountID, fgPositions.RowData(Row).SymbolOrSymbolID, False) Then
                        fgPositions.Col = GDCol(eGDCol_Source)
                    End If
                End If
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.fgPositions_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgPositions_BeforeMouseDown
'' Description: Handle the user pressing a mouse button in the grid
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Mouse Location, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgPositions_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Mouse row in the grid
    Dim lMouseCol As Long               ' Mouse column in the grid
    Dim bCheckedCell As Boolean         ' Is the cell currently checked?
    Dim strAction As String             ' Action being taken by the user
    Dim lIndex As Long                  ' Index into a for loop
    
    With fgPositions
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If (Button = vbLeftButton) And (Shift = vbCtrlMask) Then
            If (lMouseCol = GDCol(eGDCol_Enable)) And (.TextMatrix(lMouseRow, GDCol(eGDCol_Symbol)) = "Strategy") Then
                bCheckedCell = CheckedCell(fgPositions, lMouseRow, GDCol(eGDCol_Enable))
                If bCheckedCell Then
                    strAction = "ON"
                Else
                    strAction = "OFF"
                End If
                
                If InfBox("You are about to turn all of the|automated trading items " & strAction & ".||Do you want to continue?|", "?", "+Yes|-No", "Confirmation") = "Y" Then
                    Cancel = True
                    
                    If .MergeRow(.FixedRows) = True Then
                        CheckedCell(fgPositions, .FixedRows, GDCol(eGDCol_Enable)) = bCheckedCell
                    End If
                    
                    ToggleAllAutoTradeItems bCheckedCell
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.fgPositions_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgPositions_CellButtonClick
'' Description: User wants to reassign fills to change the position information
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgPositions_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim bReload As Boolean              ' Reload the grid?
    Dim AcctPos As cAccountPosition     ' Get the account position out of the grid
    Dim TradeItem As cAutoTradeItem     ' Automated trading item
    
    If TypeOf fgPositions.RowData(Row) Is cAccountPosition Then
        Set AcctPos = fgPositions.RowData(Row)
        If AcctPos.AutoTradeItemID > 0& Then
            Set TradeItem = g.TradingItems.Item(Str(AcctPos.AutoTradeItemID))
            If Not TradeItem Is Nothing Then
                frmReassignFills.ShowMe TradeItem, bReload
                
                If bReload Then
                    LoadGrid
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.fgPositions_CellButtonClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgPositions_Click
'' Description: Handle the user's click on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgPositions_Click()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.fgPositions_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgPositions_ComboCloseUp
'' Description: When the user closes the combo, finish the edit
'' Inputs:      Row and Column of edit, Whether to Finish the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgPositions_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error GoTo ErrSection:

    FinishEdit = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.fgPositions_ComboCloseUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgPositions_MouseMove
'' Description: Notification that the user has moved the mouse in the grid
'' Inputs:      Mouse Button pressed, Shift/Ctrl/Alt status, X Location of the mouse, Y Location of the mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgPositions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Static lMouseRow As Long            ' Row in the grid that the mouse is over
    Static lMouseCol As Long            ' Column in the grid that the mouse is over
    
    With fgPositions
        If (.MouseRow <> lMouseRow) Or (.MouseCol <> lMouseCol) Then
            lMouseRow = .MouseRow
            lMouseCol = .MouseCol
            
            If (lMouseRow = .FixedRows) And (lMouseCol = GDCol(eGDCol_Enable)) And (.MergeRow(lMouseRow) = True) Then
                If CheckedCell(fgPositions, lMouseRow, lMouseCol) = True Then
                    .ToolTipText = "Click here to mark all automated trading items in this list as disabled"
                Else
                    .ToolTipText = "Click here to mark all automated trading items in this list as enabled"
                End If
            Else
                .ToolTipText = ""
            End If
        End If
    End With

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
    
    g.Styler.StyleForm Me
    
    strPlacement = GetIniFileProperty("PositionConfirm", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement, "LHTW"
    End If
    Icon = Picture16("kBlank")
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, ask if everything was OK
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value from the InfBox

    If UnloadMode <> vbFormCode Then
        Cancel = False
        If (m.bShow = True) And (Len(m.strAccount) = 0) And (Len(m.strSymbol) = 0) And (m.TradeItem Is Nothing) Then
            If Not g.Broker.IsLiveAccount(m.nBroker) Then
                m.bOK = True
                SaveResults
            Else
                strReturn = InfBox("Was all of the position information correct?", "?", "+Yes|-No|Cancel", Caption)
                Select Case UCase(strReturn)
                    Case "Y"
                        m.bOK = True
                        SaveResults
                    
                    Case "N"
                        m.bOK = False
                        g.Broker.Disconnect m.nBroker, "Did not verify positions"
                    
                    Case "C"
                        Cancel = True
                        
                End Select
            End If
        Else
            m.bOK = True
            SaveResults
        End If
        
        If (Cancel = False) And (m.bModal = True) Then
            Cancel = True
            Hide
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and size controls when the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lMinHeight As Long              ' Minimum allowable height
    Dim lMinWidth As Long               ' Minimum allowable width
    Dim bRithmicBroker As Boolean       ' Is this a Rithmic broker?
    
    lMinHeight = 3660
    lMinWidth = 6450
    
    If Not LimitFormSize(Me, lMinWidth, lMinHeight) Then
        bRithmicBroker = g.Broker.IsRithmicBroker(m.nBroker)
        
        With picRithmic
            .Move 60, ScaleHeight - .Height - 60
        End With
        
        With picPbo
            .Move ScaleWidth - .Width - 60, ScaleHeight - .Height - 60
        End With
        
        With fraButtons
            If bRithmicBroker Then
                .Move (ScaleWidth - .Width) / 2, picRithmic.Top - .Height - 60
            Else
                .Move (ScaleWidth - .Width) / 2, ScaleHeight - .Height - 60
            End If
            
            If (Len(m.strAccount) <> 0) Or (Len(m.strSymbol) <> 0) Or (Not m.TradeItem Is Nothing) Then
                cmdDisconnect.Visible = False
                cmdOK.Move (.Width - cmdOK.Width) / 2
            End If
        End With
        
        With lblConfirm
            .Move 60, fraButtons.Top - .Height - 60, ScaleWidth - 120
        End With
        
        With lblIntro
            .Move 60, 60, ScaleWidth - 120
        End With
        
        With fgPositions
            If bRithmicBroker Then
                .Move 60, lblIntro.Top + lblIntro.Height + 60, ScaleWidth - 120, ScaleHeight - picRithmic.Height - fraButtons.Height - lblIntro.Height - lblConfirm.Height - 660
            Else
                .Move 60, lblIntro.Top + lblIntro.Height + 60, ScaleWidth - 120, ScaleHeight - fraButtons.Height - lblIntro.Height - lblConfirm.Height - 600
            End If
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

    If m.bShow Then
        SetIniFileProperty "PositionConfirm", GetFormPlacement(Me), "Placement", g.strIniFile
    End If

    m.bOK = False
    m.bModal = False
    m.bShow = False
    m.nBroker = -1&
    m.strAccount = ""
    m.strSymbol = ""
    
    m.bTradeItemMode = False
    Set m.TradeItem = Nothing
    Set m.astrTradeItemSymbols = Nothing
    Set m.astrTradeItemAccounts = Nothing
    Set m.alEnableAcctPosIds = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.Form_Unload"
    
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

    With fgPositions
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .HighLight = flexHighlightNever
        .MergeCells = flexMergeFree
        .OutlineBar = flexOutlineBarSimpleLeaf
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .FixedRows = 1
        .Rows = .FixedRows
        .FixedCols = 0
        .Cols = GDCol(eGDCol_NumCols)
        
        .OutlineCol = GDCol(eGDCol_Account)
        
        .TextMatrix(0, GDCol(eGDCol_Account)) = "Account"
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_Position)) = "Position"
        .TextMatrix(0, GDCol(eGDCol_Buys)) = "Buys"
        .TextMatrix(0, GDCol(eGDCol_Sells)) = "Sells"
        .TextMatrix(0, GDCol(eGDCol_Overnight)) = "Overnight"
        .TextMatrix(0, GDCol(eGDCol_Source)) = "Auto Exit or Strategy"
        .TextMatrix(0, GDCol(eGDCol_SourceID)) = "Source ID"
        .TextMatrix(0, GDCol(eGDCol_Enable)) = "Enable"
        .TextMatrix(0, GDCol(eGDCol_SymbolError)) = "Symbol Error"
        .TextMatrix(0, GDCol(eGDCol_QuantityError)) = "Quantity Error"
        
        .ColAlignment(GDCol(eGDCol_Account)) = flexAlignLeftTop
        
        ''.ColDataType(GDCol(eGDCol_Enable)) = flexDTBoolean
        
        .ColHidden(GDCol(eGDCol_SourceID)) = True
        .ColHidden(GDCol(eGDCol_Buys)) = True
        .ColHidden(GDCol(eGDCol_Sells)) = True
        .ColHidden(GDCol(eGDCol_Overnight)) = True
        .ColHidden(GDCol(eGDCol_SymbolError)) = True
        .ColHidden(GDCol(eGDCol_QuantityError)) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid
'' Inputs:      None
'' Returns:     None
''
'' FillSummary: Account, Symbol, At ID, Buys, Sells, Net, Total, PriceSum, Entries,
''              ClosedProfit, AvgEntry, Initial Fill Price, Initial Fill Date,
''              Session Date, Last Traded, Overnight
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim BrokerInfo As cBrokerInfo       ' Broker information object
    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID for a symbol
    Dim strAutoExit As String           ' Auto exit for a manual trade
    Dim lAutoTradeItemID As Long        ' Automated trading item ID
    Dim bHasAutoTrading As Boolean      ' Does the user have automated trading items?
    Dim FillSumms As cAccountPositions  ' Collection of fill summaries
    Dim strAccount As String            ' Account number
    Dim TradeItem As cAutoTradeItem     ' Automated trading item
    Dim strSymbolError As String        ' Symbol error
    Dim strQtyError As String           ' Quantity error
    Dim bAllAutoTradeEnabled As Boolean ' Are all of the automated trading items enabled?

    bAllAutoTradeEnabled = True
    
    With fgPositions
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        If g.nColorTheme = kDarkThemeColor Then .ForeColor = vbWhite 'JM 12-16-2015: need to do here because fix form controls not yet called

        Set BrokerInfo = g.Broker.BrokerInfo(m.nBroker)
        If Not BrokerInfo Is Nothing Then
            Set FillSumms = BrokerInfo.FillSummary.MakeCopy
            If Not FillSumms Is Nothing Then
                bHasAutoTrading = False
                For lIndex = 1 To FillSumms.Count
                    strAccount = g.Broker.AccountNumberForID(FillSumms(lIndex).AccountID)
                    If (IncludeRow(strAccount, FillSumms(lIndex).Symbol, FillSumms(lIndex).AutoTradeItemID) = True) And (IsExpiredContract(FillSumms(lIndex).SymbolOrSymbolID) = False) Then
                        .Rows = .Rows + 1
                        .MergeRow(.Rows - 1) = False
                        .IsSubtotal(.Rows - 1) = True
                        
                        .RowData(.Rows - 1) = FillSumms(lIndex)
                        
                        lAutoTradeItemID = FillSumms(lIndex).AutoTradeItemID
                        If lAutoTradeItemID = -1& Then
                            .RowOutlineLevel(.Rows - 1) = 1
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Account)) = strAccount
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = FillSumms(lIndex).Symbol
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Position)) = Str(BrokerInfo.BrokerPosition(strAccount, FillSumms(lIndex).Symbol))
                            .Cell(flexcpFontBold, .Rows - 1, GDCol(eGDCol_Position)) = False
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Buys)) = Str(FillSumms(lIndex).NumBuysSnapshot)
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Sells)) = Str(FillSumms(lIndex).NumSellsSnapshot)
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Overnight)) = Str(FillSumms(lIndex).CurrentPosition)
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Source)) = ""
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_SourceID)) = Str(FillSumms(lIndex).AutoTradeItemID)
                            .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Enable)) = flexNoCheckbox
                            
                            If ValidatePosition(.Rows - 1) Then
                                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = .ForeColor
                            Else
                                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                            End If
                            
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_SymbolError)) = ""
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_QuantityError)) = ""
                            
                        ElseIf lAutoTradeItemID = 0& Then
                            .RowOutlineLevel(.Rows - 1) = 2
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Account)) = ""
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = "Manual"
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Position)) = Str(FillSumms(lIndex).CurrentPositionSnapshot)
                            .Cell(flexcpFontBold, .Rows - 1, GDCol(eGDCol_Position)) = False
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Buys)) = Str(FillSumms(lIndex).NumBuysSnapshot)
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Sells)) = Str(FillSumms(lIndex).NumSellsSnapshot)
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Overnight)) = Str(FillSumms(lIndex).CurrentPosition)
                            strAutoExit = g.OrderStrategies.ExitForAccountAndSymbol(FillSumms(lIndex).AccountID, FillSumms(lIndex).SymbolOrSymbolID)
                            If (Len(strAutoExit) = 0) Then
                                .TextMatrix(.Rows - 1, GDCol(eGDCol_Source)) = "None"
                                .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Enable)) = flexUnchecked
                            ElseIf (AllowAutoExits(FillSumms(lIndex).AccountID, FillSumms(lIndex).SymbolOrSymbolID, False, m.bShow) = False) Or (FillSumms(lIndex).HasCarPosFixEntries = True) Then
                                .TextMatrix(.Rows - 1, GDCol(eGDCol_Source)) = "None"
                                .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Enable)) = flexUnchecked
                            Else
                                .TextMatrix(.Rows - 1, GDCol(eGDCol_Source)) = strAutoExit
                                .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Enable)) = flexChecked
                            End If
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_SourceID)) = Str(FillSumms(lIndex).AutoTradeItemID)
                            .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = .Cell(flexcpForeColor, .Rows - 2, 0)
                            
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_SymbolError)) = ""
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_QuantityError)) = ""
                       
                        Else
                            bHasAutoTrading = True
                            
                            Set TradeItem = New cAutoTradeItem
                            TradeItem.Load lAutoTradeItemID, False
                            
                            strSymbolError = AutomatedSymbolError(TradeItem.AccountID, TradeItem.SymbolOrSymbolID, "Automated Trading", True)
                            strQtyError = AutomatedQuantityError(TradeItem, Str(TradeItem.QtyNextEntry), strSymbolError)
                            
                            .RowOutlineLevel(.Rows - 1) = 2
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Account)) = ""
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = "Strategy"
                            '.TextMatrix(.Rows - 1, GDCol(eGDCol_Position)) = Str(g.TradingItems.CurrentPosition(FillSumms(lIndex).AutoTradeItemID, FillSumms(lIndex).SymbolOrSymbolID, FillSumms(lIndex).AccountID))
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Position)) = Str(FillSumms(lIndex).CurrentPositionSnapshot)
                            .Cell(flexcpFontBold, .Rows - 1, GDCol(eGDCol_Position)) = True
                            .Cell(flexcpPicture, .Rows - 1, GDCol(eGDCol_Position)) = Nothing
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Buys)) = Str(FillSumms(lIndex).NumBuysSnapshot)
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Sells)) = Str(FillSumms(lIndex).NumSellsSnapshot)
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Overnight)) = Str(FillSumms(lIndex).CurrentPosition)
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Source)) = AutoTradeItemNameForID(FillSumms(lIndex).AutoTradeItemID)
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_SourceID)) = Str(FillSumms(lIndex).AutoTradeItemID)
                            
                            ' DAJ 01/17/2012: Don't uncheck the box here if it is active but in a
                            ' position mismatch because that will be handled elsewhere.  If we do it
                            ' here there is a possibility we will turn off an active auto trade item
                            ' on a "first time" mismatch...
                            ' DAJ 07/08/2014: Don't disable an automated trading item here if the
                            ' max units went to zero -- we want to allow the strategy to keep running
                            ' to get out of any position they may be in...
                            ' DAJ 11/11/2014: Mark the automated trading item as enabled if it is in
                            ' the "previously active" collection of the automated trading items
                            ' collection and it has been less than five minutes since it went disabled...
                            If AllowAutoExits(FillSumms(lIndex).AccountID, FillSumms(lIndex).SymbolOrSymbolID, False, False) = False Then
                                .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Enable)) = flexUnchecked
                            'ElseIf mSysNav.MaxUnitsForAutoTradeID(FillSumms(lIndex).AutoTradeItemID) = 0 Then
                            '    .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Enable)) = flexUnchecked
                            ElseIf (Len(strSymbolError) > 0) Or (Len(strQtyError) > 0) Then
                                .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Enable)) = flexUnchecked
                            ElseIf EnableAccountPosition(FillSumms(lIndex).AccountPositionID) = True Then
                                .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Enable)) = flexChecked
                            ElseIf g.TradingItems.Active(Str(FillSumms(lIndex).AutoTradeItemID)) Then
                                .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Enable)) = flexChecked
                            ElseIf g.TradingItems.PreviouslyActive(lAutoTradeItemID) = True Then
                                .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Enable)) = flexChecked
                            Else
                                .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Enable)) = flexUnchecked
                            End If
                            
                            If .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Enable)) = flexUnchecked Then
                                bAllAutoTradeEnabled = False
                            End If
                            
                            .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = .Cell(flexcpForeColor, .Rows - 2, 0)
                            
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_SymbolError)) = strSymbolError
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_QuantityError)) = strQtyError
                            
                            If (Len(strSymbolError) > 0) Or (Len(strQtyError) > 0) Then
                                .Cell(flexcpForeColor, .Rows - 1, GDCol(eGDCol_Account), .Rows - 1, GDCol(eGDCol_SourceID)) = vbRed
                            Else
                                .Cell(flexcpForeColor, .Rows - 1, GDCol(eGDCol_Account), .Rows - 1, GDCol(eGDCol_SourceID)) = .Cell(flexcpForeColor, .Rows - 1, GDCol(eGDCol_QuantityError))
                            End If
                        
                        End If
                    End If
                Next lIndex
            End If
        End If
        
        RemoveNonActivity
        RecalcAll
        
        ' DAJ 01/14/2016: Tim wants an easier way for the user to toggle all automated trading items.  My idea is to add
        ' a row to the very top of the grid ( if there are automated trading items in the grid ) that says "All Automated
        ' Trading Items" with the check box in the enable grid.  If all items are on, the check box will be checked to start,
        ' other wise it will be unchecked...
        If bHasAutoTrading Then
            .Rows = .Rows + 1
            .RowPosition(.Rows - 1) = .FixedRows
            .MergeRow(.FixedRows) = True
            
            .Cell(flexcpText, .FixedRows, GDCol(eGDCol_Account), .FixedRows, GDCol(eGDCol_Enable) - 1) = "All Automated Trading Items"
            If bAllAutoTradeEnabled Then
                .Cell(flexcpChecked, .FixedRows, GDCol(eGDCol_Enable)) = flexChecked
            Else
                .Cell(flexcpChecked, .FixedRows, GDCol(eGDCol_Enable)) = flexUnchecked
            End If
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With
    
    If bHasAutoTrading Then
        lblIntro.Caption = "This is a summary of your current positions based on information retrieved from the broker's servers.  Please take a moment to verify that all information here is correct.  Positions in bold can be edited."
    Else
        lblIntro.Caption = "This is a summary of your current positions based on information retrieved from the broker's servers.  Please take a moment to verify that all information here is correct."
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IncludeRow
'' Description: Do we want to include the given account and symbol?
'' Inputs:      Account, Symbol, Auto Trade Item ID
'' Returns:     True if Include, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IncludeRow(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAutoTradeItemID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value from the function
    Dim lIndex As Long                  ' Index into a for loop

    bReturn = False
    
    ' If this is a stream replay account and stream replay is not on, then don't include...
    If (g.Broker.AccountTypeForNumber(strAccount) = eTT_AccountType_SimReplay) And (g.nReplaySession = 0) Then
        bReturn = False
        
    ElseIf Not m.TradeItem Is Nothing Then
        For lIndex = 0 To m.astrTradeItemAccounts.Size - 1
            bReturn = ((m.astrTradeItemAccounts(lIndex) = strAccount) And (m.astrTradeItemSymbols(lIndex) = strSymbol))
            If bReturn = True Then
                Exit For
            End If
        Next lIndex
        
    ' Otherwise if an account or symbol was specified for the form and this follows
    ' the specification OR nothing was specified, then include the row...
    ElseIf (Len(m.strAccount) = 0) Or (m.strAccount = strAccount) Then
        If (Len(m.strSymbol) = 0) Then
            bReturn = True
        ElseIf (InStr(m.strSymbol, "-0") = 0) And (m.strSymbol = strSymbol) Then
            bReturn = True
        ElseIf (InStr(m.strSymbol, "-0") <> 0) And (BaseSymbolForSymbol(m.strSymbol) = BaseSymbolForSymbol(strSymbol)) Then
            bReturn = True
        End If
        
    End If
    
    If (bReturn = True) And (lAutoTradeItemID > 0&) Then
        bReturn = Not mSysNav.AutoTradeItemIsDeleted(lAutoTradeItemID)
    End If
    
    IncludeRow = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPositionConfirm.IncludeRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoExitList
'' Description: Get a list of the auto exits for the combo list of the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AutoExitList() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array to join together for the return string
    Dim astrStrategies As cGdArray      ' Array of auto exit strategies
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    astrReturn.Add "None"
    
    Set astrStrategies = GetExitOrderStrategies
    If Not astrStrategies Is Nothing Then
        For lIndex = 0 To astrStrategies.Size - 1
            astrReturn.Add Parse(astrStrategies(lIndex), vbTab, 1)
        Next lIndex
    End If
    
    AutoExitList = astrReturn.JoinFields("|")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPositionConfirm.AutoExitList"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidatePosition
'' Description: Validate the position for the given row
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidatePosition(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim lNet As Long                    ' Net position for the row
    Dim lBuys As Long                   ' Number of buys for the row
    Dim lSells As Long                  ' Number of sells for the row
    Dim lOvernight As Long              ' Overnight position for the row
    Dim bReturn As Boolean              ' Return value from the function
    
    bReturn = False
    
    With fgPositions
        If (lRow >= .FixedRows) And (lRow < .Rows) Then
            lNet = CLng(Val(.TextMatrix(lRow, GDCol(eGDCol_Position))))
            lBuys = CLng(Val(.TextMatrix(lRow, GDCol(eGDCol_Buys))))
            lSells = CLng(Val(.TextMatrix(lRow, GDCol(eGDCol_Sells))))
            lOvernight = CLng(Val(.TextMatrix(lRow, GDCol(eGDCol_Overnight))))
            
            If (lNet = (lBuys - lSells) + lOvernight) Then
                bReturn = True
            Else
                bReturn = False
            End If
        End If
    End With
    
    ValidatePosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPositionConfirm.ValidatePosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalcAll
'' Description: Recalculate the position and overnight position for everything
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RecalcAll()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With fgPositions
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, GDCol(eGDCol_SourceID)) = "-1" Then
                RecalcChildren lIndex
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.RecalcAll"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RecalcChildren
'' Description: Recalculate the position and overnight position for children
'' Inputs:      Parent Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RecalcChildren(ByVal lParentRow As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lFirstChild As Long             ' First child of the parent
    Dim lLastChild As Long              ' Last child of the parent
    Dim lNetPosition As Long            ' Net position of the strategies
    Dim lNetOvernight As Long           ' Net overnight position of the strategies
    Dim lPosition As Long               ' Current position
    Dim lBuys As Long                   ' Number of buys
    Dim lSells As Long                  ' Number of sells
    Dim lOvernight As Long              ' Current overnight position
    Dim lTotalPosition As Long          ' Total position for the group
    Dim lTotalOvernight As Long         ' Total overnight position for the group
    
    With fgPositions
        .Redraw = flexRDNone
        
        lFirstChild = .GetNodeRow(lParentRow, flexNTFirstChild)
        lLastChild = .GetNodeRow(lParentRow, flexNTLastChild)
        
        lNetPosition = 0&
        lNetOvernight = 0&
        lTotalPosition = CLng(Val(.TextMatrix(lParentRow, GDCol(eGDCol_Position))))
        lTotalOvernight = CLng(Val(.TextMatrix(lParentRow, GDCol(eGDCol_Overnight))))
        
        If (lFirstChild <> -1&) And (lLastChild <> -1&) Then
            ' 1) Walk through automated trading strategies, fix the overnight position
            '    if necessary, and calculate net position and net overnight position of
            '    the automated trading strategies...
            For lIndex = lFirstChild To lLastChild
                If CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_SourceID)))) > 0& Then
                    lPosition = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_Position))))
                    lBuys = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_Buys))))
                    lSells = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_Sells))))
                    
                    lOvernight = lPosition - lBuys + lSells
                    .TextMatrix(lIndex, GDCol(eGDCol_Overnight)) = Str(lOvernight)
                    
                    lNetPosition = lNetPosition + lPosition
                    lNetOvernight = lNetOvernight + lOvernight
                End If
            Next lIndex
            
            ' 2) Fix the manual position and overnight position given the net position
            '    and net overnight position of the automated trading strategies and the
            '    position and overnight position of the parent row...
            For lIndex = lFirstChild To lLastChild
                If .TextMatrix(lIndex, GDCol(eGDCol_SourceID)) = "0" Then
                    .TextMatrix(lIndex, GDCol(eGDCol_Position)) = Str(lTotalPosition - lNetPosition)
                    .TextMatrix(lIndex, GDCol(eGDCol_Overnight)) = Str(lTotalOvernight - lNetOvernight)
                    
                    Exit For
                End If
            Next lIndex
        End If
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.RecalcChildren"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveNonActivity
'' Description: Remove any parents with no children
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveNonActivity()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lParent As Long                 ' Parent row
    
    With fgPositions
        .Redraw = flexRDNone
        
        ' Pass 1: Get rid of expired future contracts and automated strategy lines that
        ' "no longer exist" and have a flat position...
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If ValidItem(lIndex) = False Then
                .RemoveItem lIndex
            End If
        Next lIndex
        
        ' Pass 2: If this is a manual trading line with no siblings (i.e. no Automated
        ' Trading Strategies assigned to this symbol) and there is no Auto Exit
        ' assigned, and the position, buys, and sells are all zero for both the
        ' manual trading line and the parent, then remove the line...
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If .RowOutlineLevel(lIndex) = 2& Then
                lParent = .GetNodeRow(lIndex, flexNTParent)
                
                If .TextMatrix(lIndex, GDCol(eGDCol_Symbol)) = "Manual" Then
                    If .GetNodeRow(lIndex, flexNTNextSibling) = -1& Then
                        If (.TextMatrix(lParent, GDCol(eGDCol_Position)) = "0") And (.TextMatrix(lParent, GDCol(eGDCol_Buys)) = "0") And (.TextMatrix(lParent, GDCol(eGDCol_Sells)) = "0") Then
                            If (.TextMatrix(lIndex, GDCol(eGDCol_Source)) = "None") And (.TextMatrix(lIndex, GDCol(eGDCol_Position)) = "0") And (.TextMatrix(lIndex, GDCol(eGDCol_Buys)) = "0") And (.TextMatrix(lIndex, GDCol(eGDCol_Sells)) = "0") Then
                                .RemoveItem lIndex
                            End If
                        End If
                    End If
                End If
            End If
        Next lIndex
        
        ' Pass3: If a parent has no children, remove the row...
        For lIndex = .Rows - 1 To .FixedRows Step -1
            If .RowOutlineLevel(lIndex) = 1& Then
                If .GetNodeRow(lIndex, flexNTFirstChild) = -1& Then
                    .RemoveItem lIndex
                End If
            End If
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.RemoveNonActivity"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveOvernightPositions
'' Description: Save the overnight position overrides to the broker info object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveOvernightPositions()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strAccount As String            ' Account number
    Dim strSymbol As String             ' Symbol
    Dim lAutoTradeItemID As Long        ' Automated Trading Item ID
    Dim lParent As Long                 ' Parent row

    With fgPositions
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, GDCol(eGDCol_SourceID)) <> "-1" Then
                lParent = .GetNodeRow(lIndex, flexNTParent)
                If lParent <> -1& Then
                    strAccount = .TextMatrix(lParent, GDCol(eGDCol_Account))
                    strSymbol = .TextMatrix(lParent, GDCol(eGDCol_Symbol))
                    lAutoTradeItemID = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_SourceID))))
                
                    g.Broker.BrokerInfo(m.nBroker).OvernightPosition(strAccount, strSymbol, lAutoTradeItemID) = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_Overnight))))
                End If
            End If
        Next lIndex
    End With
    
    If m.TradeItem Is Nothing Then
        g.Broker.BrokerInfo(m.nBroker).RebuildFillSummaries m.strAccount, m.strSymbol
    Else
        For lIndex = 0 To m.astrTradeItemAccounts.Size - 1
            g.Broker.BrokerInfo(m.nBroker).RebuildFillSummaries m.astrTradeItemAccounts(lIndex), m.astrTradeItemSymbols(lIndex)
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.SaveOvernightPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpGrid
'' Description: Dump the grid to the appropriate broker log
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpGrid()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim strLine As String               ' String to dump to the log file
    
    g.Broker.BrokerDebug m.nBroker, "Position Verification Complete (Form Shown = " & Str(m.bShow) & "): "
    With fgPositions
        If .Rows > .FixedRows Then
            For lRow = .FixedRows To .Rows - 1
                strLine = .TextMatrix(lRow, 0)
                For lCol = 1 To .Cols - 1
                    If (lCol = GDCol(eGDCol_Enable)) And (.RowOutlineLevel(lRow) = 2) Then
                        strLine = strLine & vbTab & Str(CheckedCell(fgPositions, lRow, lCol))
                    Else
                        strLine = strLine & vbTab & .TextMatrix(lRow, lCol)
                    End If
                Next lCol
                
                g.Broker.BrokerDebug m.nBroker, vbTab & strLine
            Next lRow
        Else
            g.Broker.BrokerDebug m.nBroker, vbTab & "No Activity to Verify"
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPositionConfirm.DumpGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ActivateItems
'' Description: Activate any appropriate automated strategies or auto exits
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ActivateItems()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lSourceID As Long               ' Source ID
    Dim lAccountID As Long              ' Account ID
    Dim lSymbolID As Long               ' Symbol ID
    Dim strSymbol As String             ' Symbol
    Dim lParent As Long                 ' Parent row
    Dim strStrategy As String           ' Strategy name
    Dim lActivatedItems As Long         ' Number of items activated
    Dim lPosition As Long               ' Position from the grid
    
    lActivatedItems = 1&
    With fgPositions
        For lIndex = .FixedRows To .Rows - 1
            lSourceID = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_SourceID))))
            If lSourceID = 0 Then
                lParent = .GetNodeRow(lIndex, flexNTParent)
                If lParent <> -1& Then
                    lAccountID = g.Broker.AccountIDForNumber(.TextMatrix(lParent, GDCol(eGDCol_Account)))
                    strSymbol = .TextMatrix(lParent, GDCol(eGDCol_Symbol))
                    lSymbolID = GetSymbolID(strSymbol)
                    strStrategy = ExitOrderStrategyFileFromName(.TextMatrix(lIndex, GDCol(eGDCol_Source)))
                    
                    If CheckedCell(fgPositions, lIndex, GDCol(eGDCol_Enable)) = True Then
                        lActivatedItems = lActivatedItems + 1&
                        If lSymbolID = 0& Then
                            g.OrderStrategies.ActivateExit lAccountID, strSymbol, strStrategy
                        Else
                            g.OrderStrategies.ActivateExit lAccountID, lSymbolID, strStrategy
                        End If
                    Else
                        If lSymbolID = 0& Then
                            g.OrderStrategies.DeactivateExit lAccountID, strSymbol, , "Position Confirm - Check box off"
                        Else
                            g.OrderStrategies.DeactivateExit lAccountID, lSymbolID, , "Position Confirm - Check box off"
                        End If
                    End If
                End If
            ElseIf lSourceID > 0 Then
                If g.TradingItems.Exists(Str(lSourceID)) Then
                    ' DAJ 08/10/2007: Only activate/deactivate and set the current position if
                    ' the symbol matches the currently active symbol for the auto trade item...
                    'If ConvertToTradeSymbol(g.TradingItems(Str(lSourceID)).SymbolOrSymbolID) = .RowData(lIndex).SymbolOrSymbolID Then
                    If BaseSymbolForSymbol(g.TradingItems(Str(lSourceID)).TradeSymbolOrID) = BaseSymbolForSymbol(.RowData(lIndex).SymbolOrSymbolID) Then
                        If g.TradingItems(Str(lSourceID)).AccountID = .RowData(lIndex).AccountID Then
                            ' DAJ 02/26/2015: Since we did a 'SaveOvernightPositions' before we got here, but
                            ' because we also did a 'DoBrokerTimer' which could cause us to process new fills,
                            ' we need to get the current position from the appropriate broker info object
                            ' instead of from the grid ( because the grid could be outdated already )...
                            'lPosition = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_Position))))
                            lPosition = g.TradingItems.Item(Str(lSourceID)).GetCurrentPositionFromBrokerInfo
                            If lPosition <> g.TradingItems.Item(Str(lSourceID)).CurrentPosition Then
                                g.TradingItems.Item(Str(lSourceID)).CurrentPosition = lPosition
                                'g.TradingItems.Item(Str(lSourceID)).PositionSymbolOrID = .RowData(lIndex).SymbolOrSymbolID
                                g.TradingItems.Item(Str(lSourceID)).Save
                            End If
                            
                            If CheckedCell(fgPositions, lIndex, GDCol(eGDCol_Enable)) = True Then
                                lActivatedItems = lActivatedItems + 1&
                                
                                ' DAJ 12/04/2014: Only show a message here if the position confirmation
                                ' box came up visible...
                                g.TradingItems.Enable lSourceID, True, .RowData(lIndex), m.bShow
                            Else
                                g.TradingItems.Disable lSourceID, , "Check box unchecked in position confirmation form"
                            End If
                        End If
                    End If
                End If
            End If
            
            ' 12/13/2012 DAJ: Do a 'DoEvents' every 10 times through the loop so that if the user has
            ' a lot of automated trading items ( e.g. Brady Preston ), the CPU has a chance to do some
            ' other things while we are activating the items...
            ' 10/22/2014 DAJ: Tim and I are thinking we want to do this every time through the loop
            ' now because of situations with a lot of automated trading items being activated by this
            ' method...
            'If lIndex Mod 9 = 0 Then
            'If lActivatedItems Mod 9 = 0 Then
                DoEvents
            'End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.ActivateItems"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveResults
'' Description: Save off the results if the user says that things are OK
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveResults()
On Error GoTo ErrSection:

    SaveOvernightPositions
    DumpGrid
        
    If fgPositions.Rows > fgPositions.FixedRows Then
        DoBrokerTimer
        ActivateItems
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.SaveResults"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidItem
'' Description: Determine if the given row is still valid to show
'' Inputs:      Row
'' Returns:     True if valid to show, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidItem(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lParent As Long                 ' Parent row
    Dim AcctPos As cAccountPosition     ' Account position item for the row
    Dim TradeItem As cAutoTradeItem     ' Trading item for the row
    
    bReturn = True
    With fgPositions
        If .RowOutlineLevel(lRow) = 2& Then
            lParent = .GetNodeRow(lRow, flexNTParent)
            
            If IsExpiredContract(.TextMatrix(lParent, GDCol(eGDCol_Symbol))) Then
                bReturn = False
            ElseIf (.TextMatrix(lRow, GDCol(eGDCol_Symbol)) = "Strategy") Then
                If (.TextMatrix(lRow, GDCol(eGDCol_Source)) = "No Longer Exists") Then
                    If (.TextMatrix(lRow, GDCol(eGDCol_Position)) = "0") Then
                        bReturn = False
                    End If
                Else
                    Set AcctPos = .RowData(lRow)
                    If Not AcctPos Is Nothing Then
                        Set TradeItem = g.TradingItems(Str(AcctPos.AutoTradeItemID))
                        If Not TradeItem Is Nothing Then
                            bReturn = (AcctPos.AccountPositionID = TradeItem.AccountPositionID)
                        Else
                            bReturn = False
                        End If
                    Else
                        bReturn = False
                    End If
                End If
            End If
        End If
    End With
    
    ValidItem = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPositionConfirm.ValidItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolForTradeItem
'' Description: Determine the symbol for the trade item
'' Inputs:      Trade Item
'' Returns:     Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SymbolForTradeItem(ByVal TradeItem As cAutoTradeItem) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim AcctPos As cAccountPosition     ' Account Position object

    strReturn = ""
    If Not TradeItem Is Nothing Then
        Set AcctPos = g.Broker.FillSummaryForTradeItem(TradeItem)
        If Not AcctPos Is Nothing Then
            strReturn = AcctPos.Symbol
        Else
            strReturn = GetSymbol(TradeItem.TradeSymbolOrID)
        End If
    End If
    
    SymbolForTradeItem = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPositionConfirm.SymbolForTradeItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddTradeItem
'' Description: Add the given trade item information
'' Inputs:      Trade Item
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddTradeItem(ByVal TradeItem As cAutoTradeItem)
On Error GoTo ErrSection:

    Dim lAccountPositionID As Long      ' Account Position ID

    m.astrTradeItemSymbols.Add SymbolForTradeItem(TradeItem)
    m.astrTradeItemAccounts.Add g.Broker.AccountNumberForID(TradeItem.AccountID)

    If TradeItem.AccountPositionID = 0 Then
        lAccountPositionID = g.Broker.CreateFillSummaryForAutoTrade(TradeItem)
        If lAccountPositionID <> 0 Then
            TradeItem.AccountPositionID = lAccountPositionID
            TradeItem.Save
            
            m.alEnableAcctPosIds.Add TradeItem.AccountPositionID
        End If
    Else
        m.alEnableAcctPosIds.Add TradeItem.AccountPositionID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.AddTradeItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableAccountPosition
'' Description: Enable the given account position ID?
'' Inputs:      Account Position ID
'' Returns:     True if enable, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EnableAccountPosition(ByVal lAccountPositionID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    For lIndex = 0 To m.alEnableAcctPosIds.Size - 1
        If m.alEnableAcctPosIds(lIndex) = lAccountPositionID Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    EnableAccountPosition = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPositionConfirm.EnableAccountPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToggleAllAutoTradeItem
'' Description: Enable or disable an automated trading items as appropriate
'' Inputs:      Row, Set to Checked?, Show Message to User?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ToggleAutoTradeItem(ByVal lRow As Long, ByVal bChecked As Boolean, Optional ByVal bShowMessage As Boolean = True)
On Error GoTo ErrSection:

    Dim strSymbolError As String        ' Symbol error
    Dim strQtyError As String           ' Quantity error

    With fgPositions
        If bChecked = True Then
            strSymbolError = .TextMatrix(lRow, GDCol(eGDCol_SymbolError))
            strQtyError = .TextMatrix(lRow, GDCol(eGDCol_QuantityError))
            
            If Len(strSymbolError) > 0 Then
                bChecked = False
                InfBox strSymbolError, "!", , "Error"
            ElseIf Len(strQtyError) > 0 Then
                bChecked = False
                InfBox strQtyError, "!", , "Error"
            ElseIf mSysNav.MaxUnitsForAutoTradeID(.RowData(lRow).AutoTradeItemID) = 0 Then
                bChecked = False
            ElseIf CanActivateAutomatedItem(.RowData(lRow).AccountID, .RowData(lRow).SymbolOrSymbolID, "Automated Trading Item", "Position Confirm", bShowMessage) = False Then
                bChecked = False
            End If
        End If
        
        CheckedCell(fgPositions, lRow, GDCol(eGDCol_Enable)) = bChecked
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.ToggleAutoTradeItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToggleAllAutoTradeItems
'' Description: Enable or disable all automated trading items as appropriate
'' Inputs:      Set to Checked?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ToggleAllAutoTradeItems(ByVal bChecked As Boolean)
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop

    With fgPositions
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, GDCol(eGDCol_Symbol)) = "Strategy" Then
                ToggleAutoTradeItem lIndex, bChecked, False
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmPositionConfirm.ToggleAllAutoTradeItems"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllAutoTradeItemsAre
'' Description: Determine if all automated trading items are enabled or disabled
'' Inputs:      Enabled?
'' Returns:     True if all enabled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AllAutoTradeItemsAre(ByVal bEnabled As Boolean) As Boolean
On Error GoTo ErrSection:
    
    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop

    bReturn = True
    With fgPositions
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, GDCol(eGDCol_Symbol)) = "Strategy" Then
                If CheckedCell(fgPositions, lIndex, GDCol(eGDCol_Enable)) = Not bEnabled Then
                    bReturn = False
                    Exit For
                End If
            End If
        Next lIndex
    End With
    
    AllAutoTradeItemsAre = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmPositionConfirm.AllAutoTradeItemsAre"
    
End Function

