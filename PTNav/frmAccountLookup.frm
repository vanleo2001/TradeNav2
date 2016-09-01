VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmAccountLookup 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraLookup 
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
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
      Caption         =   "frmAccountLookup.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAccountLookup.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAccountLookup.frx":0040
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdFind 
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   600
         Width           =   675
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
         Caption         =   "frmAccountLookup.frx":005C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAccountLookup.frx":0084
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAccountLookup.frx":00A4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txtFind 
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   660
         Width           =   3615
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmAccountLookup.frx":00C0
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
         Alignment       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmAccountLookup.frx":00E0
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAccountLookup.frx":0100
      End
      Begin HexUniControls.ctlUniRadioXP optContains 
         Height          =   220
         Left            =   4320
         TabIndex        =   3
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAccountLookup.frx":011C
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAccountLookup.frx":014E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAccountLookup.frx":016E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optCustomer 
         Height          =   220
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAccountLookup.frx":018A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAccountLookup.frx":01D4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAccountLookup.frx":01F4
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAccount 
         Height          =   220
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAccountLookup.frx":0210
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmAccountLookup.frx":0258
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmAccountLookup.frx":0278
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblFind 
         Height          =   195
         Left            =   240
         Top             =   690
         Width           =   435
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
         Caption         =   "frmAccountLookup.frx":0294
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmAccountLookup.frx":02C0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmAccountLookup.frx":02E0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgAccounts 
      Height          =   2595
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2895
      _cx             =   5106
      _cy             =   4577
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
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   3915
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
      Caption         =   "frmAccountLookup.frx":02FC
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAccountLookup.frx":0328
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAccountLookup.frx":0348
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
         Height          =   495
         Left            =   2700
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
         Caption         =   "frmAccountLookup.frx":0364
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAccountLookup.frx":038C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAccountLookup.frx":03AC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   1380
         TabIndex        =   10
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
         Caption         =   "frmAccountLookup.frx":03C8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAccountLookup.frx":03F6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAccountLookup.frx":0416
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   0
         TabIndex        =   9
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
         Caption         =   "frmAccountLookup.frx":0432
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAccountLookup.frx":0458
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAccountLookup.frx":0478
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmAccountLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmAccountLookup.frm
'' Description: Allows the user to quickly look up an account number
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 04/25/2012   DAJ         Removed broker from account lookup call
'' 05/03/2012   DAJ         Persist last find option chosen
'' 06/05/2012   DAJ         Fix in FindRow to select first row
'' 01/30/2013   DAJ         Allow Find Button if defaults to Contains initially
'' 11/15/2013   DAJ         Allowed for default account or customer passed into ShowMe
'' 12/19/2013   DAJ         "Lauren List" tweaks
'' 01/31/2014   DAJ         Changed call for adding a new feedyard customer
'' 03/07/2014   DAJ         Moved Cattle stuff into NavCattle.DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_AccountID = 0
    eGDCol_AccountNumber
    eGDCol_CustomerName
    eGDCol_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK or Cancel?
    
    bCattle As Boolean                  ' Cattle mode?
    lSortCol As Long                    ' Last sorted column in the grid
    lSortDir As Long                    ' Last sorted direction in the grid
End Type
Private m As mPrivate

Public Function ForCattle() As Boolean
    ForCattle = m.bCattle
End Function

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Accounts, Default account number, Default customer, Cattle?
'' Returns:     Account Number (or blank if Cancelled)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal astrAccounts As cGdArray, Optional ByVal strAccountNumber As String = "", Optional ByVal strCustomer As String = "", Optional ByVal bForCattle As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strSelectedOption As String     ' Previously selected option button

    m.bCattle = bForCattle
    
    If bForCattle Then
        cmdAdd.Visible = True
        fraButtons.Width = cmdAdd.Left + cmdAdd.Width
    Else
        cmdAdd.Visible = False
        fraButtons.Width = cmdCancel.Left + cmdCancel.Width
    End If

    InitGrid
    LoadGrid astrAccounts
    
    If Len(strAccountNumber) > 0 Then
        optAccount.Value = True
        txtFind.Text = strAccountNumber
        FindRow
    ElseIf Len(strCustomer) > 0 Then
        optCustomer.Value = True
        txtFind.Text = strCustomer
        FindRow
    Else
        strSelectedOption = GetIniFileProperty("LastSelected", "optAccount", "AccountLookup", g.strIniFile)
        Select Case strSelectedOption
            Case "optAccount"
                optAccount.Value = True
            Case "optContains"
                optContains.Value = True
            Case "optCustomer"
                optCustomer.Value = True
        End Select
    End If
    
    cmdFind.Enabled = optContains.Value
    
    MoveFocus txtFind
    
    ShowForm Me, eForm_Modal, frmMain, , ALT_GRID_ROW_COLOR
    
    If m.bOK Then
        If (fgAccounts.Row >= fgAccounts.FixedRows) And (fgAccounts.Row < fgAccounts.Rows) Then
            ShowMe = fgAccounts.TextMatrix(fgAccounts.Row, GDCol(eGDCol_AccountNumber))
        End If
    End If

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmAccountLookup.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Cattle_Customer
'' Description: Handle a new cattle customer being added
'' Inputs:      Customer information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Cattle_Customer(ByVal cattleMessage As cBrokerMessage)
On Error GoTo ErrSection:

    If m.bCattle Then
        AccountToGrid cattleMessage("Number"), cattleMessage("Name"), True
        
        If optAccount.Value = True Then
            txtFind.Text = cattleMessage("Number")
        ElseIf optCustomer.Value = True Then
            txtFind.Text = cattleMessage("Name")
        End If
        
        FilterGrid
        FindRow
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.Cattle_Customer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAdd_Click
'' Description: Allow the user to add an account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAdd_Click()
On Error GoTo ErrSection:

    g.CattleBridge.UpdateFeedYardCustomer

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.cmdAdd_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Tell ShowMe to unload the form and don't return any account
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
    RaiseError "frmAccountLookup.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdFind_Click
'' Description: Filter the grid based on the user text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdFind_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    ' Filter the grid on the contains text...
    FilterGrid
    
    ' Select the first visible row...
    With fgAccounts
        For lIndex = .FixedRows To .Rows - 1
            If .RowHidden(lIndex) = False Then
                .Row = lIndex
                .RowSel = lIndex
                
                Exit For
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAccountLookup.cmdFind_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Tell ShowMe to unload the form and return the selected account
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
    RaiseError "frmAccountLookup.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccounts_AfterSort
'' Description: After sorting the grid, make sure the back colors are correct
'' Inputs:      Column sorted, Order sorted
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccounts_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SetBackColors fgAccounts
    
    m.lSortCol = Col
    m.lSortDir = Order

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.fgAccounts_AfterSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccounts_DblClick
'' Description: If the user double clicks, select the account and exit form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccounts_DblClick()
On Error GoTo ErrSection:

    fgAccounts.Row = fgAccounts.MouseRow
    fgAccounts.RowSel = fgAccounts.Row
    
    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.fgAccounts_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form is activated, move focus to the text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    MoveFocus txtFind

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize and setup the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement string from the ini file

    g.Styler.StyleForm Me
    
    Caption = "Account Lookup"
    Icon = Picture16("kBlank")
    
    strPlacement = GetIniFileProperty("frmAccountLookup", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
    End If
    
    m.lSortCol = -99999
    m.lSortDir = -99999
    
    cmdOK.Default = True
    cmdCancel.Cancel = True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAccountLookup.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Determine whether we really want to unload the form or not
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
    RaiseError "frmAccountLookup.Form_QueryUnload"
    
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

    If LimitFormSize(Me, fraLookup.Width + 120, fraLookup.Height * 3 + 240) Then Exit Sub
    
    With fraLookup
        .Move 60, 60
    End With
    
    With fgAccounts
        .Move 60, fraLookup.Height + 120, ScaleWidth - 120, ScaleHeight - fraLookup.Height - fraButtons.Height - 240
    End With
    
    With fraButtons
        .Move (ScaleWidth / 2) - (.Width / 2), ScaleHeight - .Height - 60
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up and save settings when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "frmAccountLookup", GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAccountLookup.Form_Unload"
    
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

    With fgAccounts
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .AutoSearch = flexSearchFromTop
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_AccountID)) = "Account ID"
        .TextMatrix(0, GDCol(eGDCol_AccountNumber)) = "Account"
        .TextMatrix(0, GDCol(eGDCol_CustomerName)) = "Customer"
        
        .ColHidden(GDCol(eGDCol_AccountID)) = True
        
        .ColAlignment(GDCol(eGDCol_AccountNumber)) = flexAlignLeftTop
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid with the available accounts
'' Inputs:      Accounts
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid(ByVal astrAccounts As cGdArray)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strAccount As String            ' Account number for the account

    With fgAccounts
        .Redraw = flexRDNone
        
        For lIndex = 0 To astrAccounts.Size - 1
            strAccount = Parse(astrAccounts(lIndex), vbTab, 1)
            If Len(strAccount) > 0 Then
                AccountToGrid strAccount, Parse(astrAccounts(lIndex), vbTab, 3), False
            End If
        Next lIndex
        
        SortGrid eGDCol_AccountNumber, flexSortGenericAscending
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optAccount_Click
'' Description: Allow the user to select an account based on account number
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optAccount_Click()
On Error GoTo ErrSection:

    If Visible Then
        SetIniFileProperty "LastSelected", "optAccount", "AccountLookup", g.strIniFile
        
        FilterGrid
        FindRow
        cmdFind.Enabled = False
        cmdOK.Default = True
        MoveFocus txtFind
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.optAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optContains_Click
'' Description: Allow the user to select an account based on containment
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optContains_Click()
On Error GoTo ErrSection:

    If Visible Then
        SetIniFileProperty "LastSelected", "optContains", "AccountLookup", g.strIniFile
        
        FilterGrid
        cmdFind.Enabled = True
        cmdFind.Default = True
        MoveFocus txtFind
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.optContains_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optCustomer_Click
'' Description: Allow the user to select an account based on customer name
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optCustomer_Click()
On Error GoTo ErrSection:

    If Visible Then
        SetIniFileProperty "LastSelected", "optCustomer", "AccountLookup", g.strIniFile
        
        FilterGrid
        FindRow
        cmdFind.Enabled = False
        cmdOK.Default = True
        MoveFocus txtFind
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.optCustomer_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtFind_Change
'' Description: As the text in the find box changes, try to go to that account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtFind_Change()
On Error GoTo ErrSection:

    If cmdFind.Enabled = True Then cmdFind.Default = True
    FindRow
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.txtFind_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterGrid
'' Description: Filter the grid based on the "Contains" text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With fgAccounts
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            If (Len(Trim(txtFind.Text)) = 0) Or (optContains.Value = False) Then
                .RowHidden(lIndex) = False
            ElseIf (InStr(UCase(.TextMatrix(lIndex, GDCol(eGDCol_AccountNumber))), UCase(txtFind.Text)) = 0) And (InStr(UCase(.TextMatrix(lIndex, GDCol(eGDCol_CustomerName))), UCase(txtFind.Text)) = 0) Then
                .RowHidden(lIndex) = True
            Else
                .RowHidden(lIndex) = False
            End If
        Next lIndex
        
        SetBackColors fgAccounts
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.FilterGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FindRow
'' Description: Find the row where the appropriate column starts with the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FindRow()
On Error GoTo ErrSection:

    Dim lCol As Long                    ' Column of the grid to deal with
    Dim lRow As Long                    ' Row returned from the FindRow

    If optContains.Value = False Then
        With fgAccounts
            If optAccount.Value = True Then
                lCol = GDCol(eGDCol_AccountNumber)
            ElseIf optCustomer.Value = True Then
                lCol = GDCol(eGDCol_CustomerName)
            End If
            
            SortGrid lCol, flexSortGenericAscending
            
            lRow = .FindRow(Trim(txtFind.Text), .FixedRows, lCol, False, False)
            If lRow >= .FixedRows And lRow < .Rows Then
                .Row = lRow
                .RowSel = lRow
                .ShowCell lRow, lCol
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.FindRow"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SortGrid
'' Description: Sort the grid by the given column and direction
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SortGrid(ByVal lCol As eGDCols, ByVal lDirection As SortSettings)
On Error GoTo ErrSection:

    If (lCol <> m.lSortCol) Or (lDirection <> m.lSortDir) Then
        fgAccounts.Col = lCol
        fgAccounts.Sort = lDirection
        SetBackColors fgAccounts
        m.lSortCol = lCol
        m.lSortDir = lDirection
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.SortGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountToGrid
'' Description: Add the given information to the grid
'' Inputs:      Account, Name, Check First?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AccountToGrid(ByVal strAccount As String, ByVal strName As String, Optional ByVal bCheckFirst As Boolean = True)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Row in the grid

    With fgAccounts
        nRedraw = .Redraw
        .Redraw = flexRDNone
            
        lRow = -1&
        If bCheckFirst Then
            For lIndex = .FixedRows To .Rows - 1
                If .TextMatrix(lIndex, GDCol(eGDCol_AccountNumber)) = strAccount Then
                    lRow = lIndex
                    Exit For
                End If
            Next lIndex
        End If
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .TextMatrix(lRow, GDCol(eGDCol_AccountNumber)) = strAccount
        .TextMatrix(lRow, GDCol(eGDCol_CustomerName)) = strName
            
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAccountLookup.AccountToGrid"
    
End Sub

