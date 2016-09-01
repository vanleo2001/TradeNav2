VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmTurnkeyAccounts 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1035
      Left            =   3240
      TabIndex        =   1
      Top             =   180
      Width           =   1215
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   540
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgAccounts 
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
End
Attribute VB_Name = "frmTurnkeyAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTurnkeyAccounts.frm
'' Description: Form for allowing user to setup Turnkey accounts
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 10/22/2012   DAJ         Rename Turnkey to HedgeLinc
'' 01/30/2013   DAJ         Include all accounts in grid ( hide accounts they don't have )
'' 11/15/2013   DAJ         Changed the way to get Turnkey icon for the form
'' 03/07/2014   DAJ         Moved into NavCattle.DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_On = 0
    eGDCol_AccountNumber
    eGDCol_Broker
    eGDCol_Status
    eGDCol_Key
    eGDCol_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?
End Type
Private m As mPrivate

Private Property Get GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Broker Accounts, Associated Accounts, Feedyard
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal BrokerAccounts As cGdTree, AssociatedAccounts As cGdTree, ByVal strFeedyard As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    'Caption = "Accounts to use with " & kTurnkeyCompanyName
    Caption = "Accounts to use with " & strFeedyard

    InitGrid
    LoadGrid BrokerAccounts, AssociatedAccounts

    ShowForm Me, eForm_Modal, g.frmMain
    
    If m.bOK Then
        AssociatedAccounts.Clear
        With fgAccounts
            For lIndex = .FixedRows To .Rows - 1
                If CheckedCell(fgAccounts, lIndex, GDCol(eGDCol_On)) = True Then
                    AssociatedAccounts.Add .RowData(lIndex), .TextMatrix(lIndex, GDCol(eGDCol_Key))
                End If
            Next lIndex
        End With
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTurnkeyAccounts.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: The user chose to cancel the dialog
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
    RaiseError "frmTurnkeyAccounts.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: The user chose to OK the dialog
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
    RaiseError "frmTurnkeyAccounts.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccounts_Click
'' Description: Handle the user toggling the "On" column
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccounts_Click()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row for the cell that the user clicked on
    Dim lMouseCol As Long               ' Column for the cell that the user clicked on
    
    With fgAccounts
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
            If lMouseCol = GDCol(eGDCol_On) Then
                CheckedCell(fgAccounts, lMouseRow, lMouseCol) = Not CheckedCell(fgAccounts, lMouseRow, lMouseCol)
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTurnkeyAccounts.fgAccounts_Click"
    
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

    Icon = g.AppBridge.Picture16(g.Turnkey.IconName)
    PlaceForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTurnkeyAccounts.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Allow ShowMe to close the dialog when the user clicks on the 'X'
'' Inputs:      Cancel the unload?, Mode of the Unload
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
    RaiseError "frmTurnkeyAccounts.Form_QueryUnload"
    
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
    Dim lSpace As Long                  ' Space between controls
    
    lSpace = 60
    lMinScaleWidth = (fraButtons.Width * 3) + (lSpace * 3)
    lMinScaleHeight = fraButtons.Height + (lSpace * 2)
    
    If Not LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) Then
        With fraButtons
            .Move ScaleWidth - .Width - lSpace, lSpace
        End With
        
        With fgAccounts
            .Move lSpace, lSpace, ScaleWidth - fraButtons.Width - (lSpace * 3), ScaleHeight - (lSpace * 2)
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

    SaveFormPlacement Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTurnkeyAccounts.Form_Unload"
    
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

    Dim nRedraw As RedrawSettings       ' State of the grid's redraw setting

    With fgAccounts
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = True
        .MergeCells = flexMergeNever
        .OutlineBar = flexOutlineBarNone
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        .Cols = GDCol(eGDCol_NumCols)
        
        .TextMatrix(0, GDCol(eGDCol_On)) = "Use"
        .TextMatrix(0, GDCol(eGDCol_AccountNumber)) = "Account"
        .TextMatrix(0, GDCol(eGDCol_Broker)) = "Broker"
        .TextMatrix(0, GDCol(eGDCol_Status)) = "Status"
        .TextMatrix(0, GDCol(eGDCol_Key)) = "Key"
        
        .ColHidden(GDCol(eGDCol_Key)) = True
        
        .ColAlignment(GDCol(eGDCol_AccountNumber)) = flexAlignLeftTop
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTurnkeyAccounts.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid
'' Inputs:      Broker Accounts, Associated Accounts
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid(ByVal BrokerAccounts As cGdTree, ByVal AssociatedAccounts As cGdTree)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' State of the grid's redraw setting
    Dim lIndex As Long                  ' Index into a for loop
    Dim Account As cBrokerMessage       ' Broker account
    Dim strKey As String                ' Key into the collection
    Dim AccountMap As cGdTree           ' Account to Row map
    Dim lRow As Long                    ' Row in the grid
    
    With fgAccounts
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        Set AccountMap = New cGdTree
        
        .Rows = .FixedRows
        For lIndex = 1 To BrokerAccounts.Count
            Set Account = BrokerAccounts(lIndex)
            strKey = BrokerAccounts.Key(lIndex)
            
            .Rows = .Rows + 1
            lRow = .Rows - 1
            
            AccountToGrid Account, strKey, lRow
            CheckedCell(fgAccounts, lRow, GDCol(eGDCol_On)) = False
            AccountMap.Add lRow, strKey
        Next lIndex
        
        For lIndex = 1 To AssociatedAccounts.Count
            Set Account = AssociatedAccounts(lIndex)
            
            strKey = Account("Broker") & "|" & Account("Number")
            If AccountMap.Exists(strKey) Then
                lRow = AccountMap(strKey)
                
                .RowData(lRow) = Account
            Else
                .Rows = .Rows + 1
                lRow = .Rows - 1
                
                AccountToGrid Account, strKey, lRow
                .RowHidden(lRow) = True
            End If
        
            CheckedCell(fgAccounts, lRow, GDCol(eGDCol_On)) = True
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTurnkeyAccounts.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountToGrid
'' Description: Add a row to the grid for the given account
'' Inputs:      Account, Key, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AccountToGrid(ByVal Account As cBrokerMessage, ByVal strKey As String, ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgAccounts
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .RowData(lRow) = Account
        
        If Len(Account("FcmNumber")) > 0 Then
            .TextMatrix(lRow, GDCol(eGDCol_AccountNumber)) = Account("FcmNumber")
        Else
            .TextMatrix(lRow, GDCol(eGDCol_AccountNumber)) = Account("Number")
        End If
        .TextMatrix(lRow, GDCol(eGDCol_Broker)) = g.AppBridge.BrokerName(CLng(Val(Account("Broker"))))
        .TextMatrix(lRow, GDCol(eGDCol_Status)) = g.BrokerEnums.ConnectionStatusToString(g.AppBridge.ConnectionStatusForAccount(Account("Number"), True))
        .TextMatrix(lRow, GDCol(eGDCol_Key)) = strKey
    
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTurnkeyAccounts.AccountToGrid"
    
End Sub
