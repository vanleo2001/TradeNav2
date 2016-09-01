VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmAccountInfo 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2895
      Left            =   3180
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   540
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgAccountInfo 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _cx             =   5106
      _cy             =   5106
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
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmAccountInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmAccountInfo.frm
'' Description: Allows the user to view advanced account information from the broker
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 08/25/2009   DAJ         Added support for PFG account information
'' 10/16/2013   DAJ         Removed PFG and Xpress
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strAccount As String                ' Account number
    strAccountInfo As String            ' Account information in a string
    
    PfgAccountInfo As cPfgAccountInfo
End Type
Private m As mPrivate

Public Property Get Account()
    Account = m.strAccount
End Property

Public Property Get AccountInfo() As String
    AccountInfo = m.strAccountInfo
End Property
Public Property Let AccountInfo(ByVal strAccountInfo As String)
    m.strAccountInfo = strAccountInfo
    LoadGridXpress
End Property

Public Property Get PfgAccountInfo() As cPfgAccountInfo
    PfgAccountInfo = m.PfgAccountInfo
End Property
Public Property Let PfgAccountInfo(ByVal AccountInfo As cPfgAccountInfo)
    Set m.PfgAccountInfo = AccountInfo
    LoadGridPfg
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Set up and show the form
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal vAccountNumberOrID As Variant)
On Error GoTo ErrSection:

    Dim nBroker As eTT_AccountType      ' Broker for the given account

    m.strAccount = g.Broker.GetAccountNumber(vAccountNumberOrID)
    Caption = "Account Information for " & m.strAccount
    
    nBroker = g.Broker.AccountTypeForNumber(m.strAccount)
    
    InitGrid nBroker
    If g.Broker.ConnectionStatusForAccount(vAccountNumberOrID) = eGDConnectionStatus_Connected Then
        RefreshAccountInfo
    Else
        m.strAccountInfo = GetIniFileProperty(m.strAccount, "", "AccountInfo", g.strIniFile)
        LoadGrid nBroker
    End If
    
    ShowForm Me, eForm_Nonmodal, frmMain

ErrExit:
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmAccountInfo.ShowMe"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow the user to print the account information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintMe()
On Error GoTo ErrSection:

    frmPrintPreview.ShowMe "TN AccountInfo", Me, 0
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.PrintMe"
            
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function for the print preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        .FontName = "Times New Roman"
        .FontSize = 14
        .FontBold = True
        .TextAlign = taCenterMiddle
        .Text = "Account Information for " & m.strAccount
        .TextAlign = taLeftMiddle
        .FontBold = False
        
        .Paragraph = ""
        .Paragraph = ""
        
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable fgAccountInfo
        Else
            .RenderControl = fgAccountInfo.hWnd
        End If
        
        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.GenerateReport"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClose_Click
'' Description: Close the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClose_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.cmdClose_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPrint_Click
'' Description: Allow the user to print the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPrint_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.cmdPrint_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRefresh_Click
'' Description: Allow the user to ask for a refreh of the information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRefresh_Click()
On Error GoTo ErrSection:

    RefreshAccountInfo

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.cmdRefresh_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccountInfo_BeforeMouseDown
'' Description: Show the popup menu if the user right clicks
'' Inputs:      Button pressed, Shift/Ctrl/Alt status, Location of Mouse,
''              Whether to Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccountInfo_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.fgAccountInfo_BeforeMouseDown"

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

    Dim strPlacement As String          ' Placement of the form from the INI file
    Dim strFont As String               ' Font information from the ini file

    Icon = Picture16(ToolbarIcon("ID_TradeTracker"))
    
    strPlacement = GetIniFileProperty("AccountInfo", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement, "LHTW"
    End If

    strFont = GetIniFileProperty("AccountInfo", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString fgAccountInfo.Font, strFont
    
    mnuPopUp.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.Form_Load"
    
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

    If LimitFormSize(Me, fraButtons.Width * 4 + 170, fraButtons.Height + 120) = False Then
        With fraButtons
            .Move ScaleWidth - 60 - .Width, 60
        End With
        
        With fgAccountInfo
            .Move 60, 60, fraButtons.Left - 120, ScaleHeight - 120
        
            If .Cols = 4 Then
                .ColWidth(0) = .Width / 3
                .ColWidth(1) = .Width / 6
                .ColWidth(2) = .Width / 3
                '.ColWidth(3) = .Width / 6
            ElseIf .Cols = 5 Then
                .ColWidth(0) = .Width / .Cols
                .ColWidth(1) = .Width / .Cols
                .ColWidth(2) = .Width / .Cols
                .ColWidth(3) = .Width / .Cols
                '.ColWidth(4) = .Width / 5
            End If
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up upon the form getting unloaded
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "AccountInfo", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "AccountInfo", FontToString(fgAccountInfo.Font), "Fonts", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.Form_Unload"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Allow the user to change the font on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgAccountInfo, False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.mnuChangeFont_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPrint_Click
'' Description: Allow the user to print the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPrint_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.mnuPrint_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuRefresh_Click
'' Description: Allow the user to refresh the data
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRefresh_Click()
On Error GoTo ErrSection:

    RefreshAccountInfo

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.mnuRefresh_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid appropriate to the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

'    Select Case nBroker
'        Case eTT_AccountType_LindWaldock
'            InitGridXpress
'        Case eTT_AccountType_ManExpress
'            InitGridXpress
'        Case eTT_AccountType_PFG
'            InitGridPfg
'    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGridXpress
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGridXpress()
On Error GoTo ErrSection:

    With fgAccountInfo
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorBkg = .BackColor
        .Editable = flexEDNone
        .ExtendLastCol = True
        .GridLines = flexGridNone
        .MergeCells = flexMergeFree
        .ScrollTrack = True
        .SelectionMode = flexSelectionFree
        .SheetBorder = .BackColor
        
        .Rows = 27
        .FixedRows = 0
        .Cols = 5
        .FixedCols = 0
        
        .Cell(flexcpText, 0, 0, 0, .Cols - 1) = "Account Information"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .MergeRow(0) = True
        
        .TextMatrix(1, 0) = "Account:"
        .TextMatrix(1, 3) = "Description:"
        .TextMatrix(2, 0) = "Current Balance:"
        .TextMatrix(2, 3) = "Last Activity Date:"
        .TextMatrix(3, 0) = "Margin Deficit Excess:"
        .TextMatrix(3, 3) = "MM Total Equity:"
        .TextMatrix(4, 0) = "Net Change on Positions:"
        .TextMatrix(4, 3) = "Positions and Orders:"
        .TextMatrix(5, 0) = "Purchasing Power:"
        .TextMatrix(5, 3) = "Securities on Deposit:"
        .TextMatrix(6, 0) = "Regulatory Code:"
        
        .Cell(flexcpText, 8, 0, 8, .Cols - 1) = "Margin Requirements"
        .Cell(flexcpAlignment, 8, 0, 8, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontBold, 8, 0, 8, .Cols - 1) = True
        .MergeRow(8) = True
        
        .TextMatrix(9, 0) = "Initial Fills:"
        .TextMatrix(9, 3) = "Initial Orders:"
        .TextMatrix(10, 0) = "Maintenance Fills:"
        .TextMatrix(10, 3) = "Maintenance Orders:"
        
        .Cell(flexcpText, 12, 0, 12, .Cols - 1) = "Option Valuation"
        .Cell(flexcpAlignment, 12, 0, 12, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontBold, 12, 0, 12, .Cols - 1) = True
        .MergeRow(12) = True
        
        .TextMatrix(13, 0) = "Long Option Value:"
        .TextMatrix(13, 3) = "Short Option Value:"
        .TextMatrix(14, 0) = "Option Premium:"
        
        .Cell(flexcpText, 16, 0, 16, .Cols - 1) = "Start of Day"
        .Cell(flexcpAlignment, 16, 0, 16, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontBold, 16, 0, 16, .Cols - 1) = True
        .MergeRow(16) = True
        
        .TextMatrix(17, 0) = "Balance:"
        .TextMatrix(17, 3) = "Open Trade Equity:"
        .TextMatrix(17, 0) = "Initial Margin:"
        .TextMatrix(17, 3) = "Maintenance Margin:"
        
        .Cell(flexcpText, 20, 0, 20, .Cols - 1) = "Commissions"
        .Cell(flexcpAlignment, 20, 0, 20, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontBold, 20, 0, 20, .Cols - 1) = True
        .MergeRow(20) = True
        
        .TextMatrix(21, 0) = "Month:"
        .TextMatrix(21, 3) = "Year:"
        
        .Cell(flexcpText, 23, 0, 23, .Cols - 1) = "Currency"
        .Cell(flexcpAlignment, 23, 0, 23, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontBold, 23, 0, 23, .Cols - 1) = True
        .MergeRow(23) = True
        
        .TextMatrix(24, 0) = "Currency:"
        .TextMatrix(24, 3) = "Description:"
        .TextMatrix(25, 0) = "Exchange Rate:"
        
        .ColAlignment(0) = flexAlignLeftTop
        .ColAlignment(1) = flexAlignRightTop
        .ColAlignment(3) = flexAlignLeftTop
        .ColAlignment(4) = flexAlignRightTop
        
        .ColWidth(0) = .Width / .Cols
        .ColWidth(1) = .Width / .Cols
        .ColWidth(2) = .Width / .Cols
        .ColWidth(3) = .Width / .Cols
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.InitGridXpress"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGridPfg
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGridPfg()
On Error GoTo ErrSection:

    With fgAccountInfo
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorBkg = .BackColor
        .Editable = flexEDNone
        .ExtendLastCol = True
        .GridLines = flexGridNone
        .MergeCells = flexMergeFree
        .ScrollTrack = True
        .SelectionMode = flexSelectionFree
        .SheetBorder = .BackColor
        
        .Rows = 18
        .FixedRows = 0
        .Cols = 4
        .FixedCols = 0
        
        .Cell(flexcpText, 0, 0, 0, .Cols - 1) = "Account Information"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .MergeRow(0) = True
        
        .TextMatrix(1, 0) = "Cash Value:"
        .TextMatrix(1, 2) = "Liquid Value:"
        .TextMatrix(2, 0) = "Beginning Cash:"
        .TextMatrix(2, 2) = "Securities on Deposit:"
        .TextMatrix(3, 0) = "Overnight Market Value:"
        .TextMatrix(3, 2) = "Cash Balance:"
        .TextMatrix(4, 0) = "Available Equity:"
        .TextMatrix(4, 2) = "Overnight Equity:"
        .TextMatrix(5, 0) = "Opening Excess Equity:"
        .TextMatrix(5, 2) = "Scalped Profit:"
        
        .Cell(flexcpText, 7, 0, 7, .Cols - 1) = "Margin Information"
        .Cell(flexcpAlignment, 7, 0, 7, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontBold, 7, 0, 7, .Cols - 1) = True
        .MergeRow(7) = True
        
        .TextMatrix(8, 0) = "Margin Required:"
        .TextMatrix(8, 2) = "Margin Excess:"
        .TextMatrix(9, 0) = "Opening Margin:"
        .TextMatrix(9, 2) = "Intraday Margin:"
        .TextMatrix(10, 0) = "Maintenance Margin:"
        .TextMatrix(10, 2) = "Maintenance Margin Excess:"
        
        .Cell(flexcpText, 12, 0, 12, .Cols - 1) = "Profit/Loss Information"
        .Cell(flexcpAlignment, 12, 0, 12, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontBold, 12, 0, 12, .Cols - 1) = True
        .MergeRow(12) = True
        
        .TextMatrix(13, 0) = "Open Pnl:"
        .TextMatrix(13, 2) = "Closed Pnl:"
        .TextMatrix(14, 0) = "Net Pnl:"
        
        .Cell(flexcpText, 16, 0, 16, .Cols - 1) = "Option Valuation"
        .Cell(flexcpAlignment, 16, 0, 16, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpFontBold, 16, 0, 16, .Cols - 1) = True
        .MergeRow(16) = True
        
        .TextMatrix(17, 0) = "LOV/SOV:"
               
        .ColAlignment(0) = flexAlignLeftTop
        .ColAlignment(1) = flexAlignRightTop
        .ColAlignment(2) = flexAlignLeftTop
        .ColAlignment(3) = flexAlignRightTop
        
        .ColWidth(0) = .Width / 3
        .ColWidth(1) = .Width / 6
        .ColWidth(2) = .Width / 3
        '.ColWidth(3) = .Width / 6
    
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.InitGridPfg"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid appropriate to the given broker
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

'    Select Case nBroker
'        Case eTT_AccountType_LindWaldock
'            LoadGridXpress
'        Case eTT_AccountType_ManExpress
'            LoadGridXpress
'        Case eTT_AccountType_PFG
'            LoadGridPfg
'    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridXpress
'' Description: Load the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridXpress()
On Error GoTo ErrSection:

    Dim astrInfo As cGdArray            ' Array of split out account information
    
    Set astrInfo = New cGdArray
    astrInfo.Create eGDARRAY_Strings

    If Len(m.strAccountInfo) > 0 Then
        astrInfo.SplitFields m.strAccountInfo, vbTab
        
        With fgAccountInfo
            .Redraw = flexRDNone
            
            .TextMatrix(1, 1) = astrInfo(0)
            .TextMatrix(1, 4) = astrInfo(1)
            .TextMatrix(2, 1) = astrInfo(10)
            If Len(astrInfo(14)) >= 8 Then
                .TextMatrix(2, 4) = Left(astrInfo(14), 8)
            Else
                .TextMatrix(2, 4) = astrInfo(14)
            End If
            .TextMatrix(3, 1) = astrInfo(16)
            .TextMatrix(3, 4) = astrInfo(17)
            .TextMatrix(4, 1) = astrInfo(17)
            .TextMatrix(4, 4) = astrInfo(22)
            .TextMatrix(5, 1) = astrInfo(23)
            .TextMatrix(5, 4) = astrInfo(25)
            .TextMatrix(6, 1) = astrInfo(24)
            
            .TextMatrix(9, 1) = astrInfo(12)
            .TextMatrix(9, 4) = astrInfo(20)
            .TextMatrix(10, 1) = astrInfo(13)
            .TextMatrix(10, 4) = astrInfo(21)
            
            .TextMatrix(13, 1) = astrInfo(15)
            .TextMatrix(13, 4) = astrInfo(26)
            .TextMatrix(14, 1) = astrInfo(19)
            
            .TextMatrix(17, 1) = astrInfo(2)
            .TextMatrix(17, 4) = astrInfo(5)
            .TextMatrix(17, 1) = astrInfo(3)
            .TextMatrix(17, 4) = astrInfo(4)
            
            .TextMatrix(21, 1) = astrInfo(6)
            .TextMatrix(21, 4) = astrInfo(7)
            
            .TextMatrix(24, 1) = astrInfo(8)
            .TextMatrix(24, 4) = astrInfo(9)
            .TextMatrix(25, 1) = astrInfo(11)
            
            .Redraw = flexRDBuffered
        End With
        
        SetIniFileProperty m.strAccount, m.strAccountInfo, "AccountInfo", g.strIniFile
        InfBox ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.LoadGridXpress"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridPfg
'' Description: Load the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridPfg()
On Error GoTo ErrSection:

    If m.PfgAccountInfo Is Nothing Then
        If Len(m.strAccountInfo) > 0 Then
            Set m.PfgAccountInfo = New cPfgAccountInfo
            m.PfgAccountInfo.FromString m.strAccountInfo
        End If
    End If
    
    If Not m.PfgAccountInfo Is Nothing Then
        With fgAccountInfo
            .Redraw = flexRDNone
            
            SetCell 1, 1, m.PfgAccountInfo.CashValue
            SetCell 1, 3, m.PfgAccountInfo.LiquidValue
            SetCell 2, 1, m.PfgAccountInfo.BeginningCash
            SetCell 2, 3, m.PfgAccountInfo.SecuritiesOnDeposit
            SetCell 3, 1, m.PfgAccountInfo.OvernightMarketValue
            SetCell 3, 3, m.PfgAccountInfo.CashBalance
            SetCell 4, 1, m.PfgAccountInfo.AvailableEquity
            SetCell 4, 3, m.PfgAccountInfo.OvernightEquity
            SetCell 5, 1, m.PfgAccountInfo.OpeningExcessEquity
            SetCell 5, 3, m.PfgAccountInfo.ScalpedProfit
            
            SetCell 8, 1, m.PfgAccountInfo.MarginRequired
            SetCell 8, 3, m.PfgAccountInfo.MarginExcess
            SetCell 9, 1, m.PfgAccountInfo.OpeningMargin
            SetCell 9, 3, m.PfgAccountInfo.IntradayMargin
            SetCell 10, 1, m.PfgAccountInfo.MaintenanceMargin
            SetCell 10, 3, m.PfgAccountInfo.MaintenanceMarginExcess
            
            SetCell 13, 1, m.PfgAccountInfo.OpenPnl
            SetCell 13, 3, m.PfgAccountInfo.ClosedPnl
            SetCell 14, 1, m.PfgAccountInfo.NetPnl
            
            SetCell 17, 1, m.PfgAccountInfo.LovSov
            
            .Redraw = flexRDBuffered
        End With
            
        SetIniFileProperty m.strAccount, m.PfgAccountInfo.ToString, "AccountInfo", g.strIniFile
    End If
    
    InfBox ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.LoadGridPfg"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshAccountInfo
'' Description: Refresh the account information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshAccountInfo()
On Error GoTo ErrSection:

    If g.Broker.ConnectionStatusForAccount(m.strAccount) = eGDConnectionStatus_Connected Then
        InfBox "Requesting account information for " & m.strAccount, , , "Account Information Request", True
        g.Broker.RefreshAccountInfo m.strAccount
    Else
        InfBox "You cannot refresh account information for '" & m.strAccount & "' because you are not currently connected to that account.", "!", , "Account Information Request"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.RefreshAccountInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCell
'' Description: Set the given cell with the given value and color
'' Inputs:      Row, Column, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCell(ByVal lRow, ByVal lCol, ByVal dValue)
On Error GoTo ErrSection:

    With fgAccountInfo
        .TextMatrix(lRow, lCol) = Format(dValue, "$#,##0.00")
        
        If dValue > 0 Then
            .Cell(flexcpForeColor, lRow, lCol) = QBColor(2)
        ElseIf dValue = 0 Then
            .Cell(flexcpForeColor, lRow, lCol) = .Cell(flexcpForeColor, lRow, 0)
        Else
            .Cell(flexcpForeColor, lRow, lCol) = vbRed
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountInfo.SetCell"

End Sub
