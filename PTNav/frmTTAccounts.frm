VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTTAccounts 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7065
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   2775
      Left            =   3360
      TabIndex        =   1
      Top             =   180
      Width           =   1395
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
      Caption         =   "frmTTAccounts.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTTAccounts.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTTAccounts.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdPrint 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   1800
         Width           =   1395
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
         Caption         =   "frmTTAccounts.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTAccounts.frx":00A6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTAccounts.frx":00C6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdView 
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1395
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
         Caption         =   "frmTTAccounts.frx":00E2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTAccounts.frx":011C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTAccounts.frx":013C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdExit 
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   2400
         Width           =   1395
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
         Caption         =   "frmTTAccounts.frx":0158
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTAccounts.frx":0182
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTAccounts.frx":01A2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   1260
         Width           =   1395
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
         Caption         =   "frmTTAccounts.frx":01BE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTAccounts.frx":01FC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTAccounts.frx":021C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   840
         Width           =   1395
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
         Caption         =   "frmTTAccounts.frx":0238
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTAccounts.frx":0272
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTAccounts.frx":0292
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNew 
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   420
         Width           =   1395
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
         Caption         =   "frmTTAccounts.frx":02AE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTTAccounts.frx":02E6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTTAccounts.frx":0306
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgAccounts 
      Height          =   2895
      Left            =   180
      TabIndex        =   0
      Top             =   180
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
      Begin VB.Menu mnuViewPositions 
         Caption         =   "View Account History"
      End
      Begin VB.Menu mnuEditAccount 
         Caption         =   "Edit Account"
      End
      Begin VB.Menu mnuNewAccount 
         Caption         =   "New Account"
      End
      Begin VB.Menu mnuDeleteAccount 
         Caption         =   "Delete Account"
      End
      Begin VB.Menu mnuPrintAccounts 
         Caption         =   "Print Accounts"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmTTAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTTAccounts.frm
'' Description: Form to allow the user to select the account they wish to edit
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 07/25/2002   DAJ         Created
'' 05/22/2009   DAJ         Changed the way that a Delete happens
'' 05/27/2009   DAJ         Hide/Show menu items when buttons are hidden/shown
'' 06/24/2011   DAJ         Remove account from grid upon delete
'' 07/28/2011   DAJ         Show New button even if called from Trade Tracker
'' 09/10/2014   DAJ         Consolidate the delete account code
'' 01/28/2016   DAJ         Fix for being allowed to delete the last sim account from here
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_AccountID = 0
    eGDCol_Number
    eGDCol_Name
    eGDCol_StartDate
    eGDCol_StartBalance
    eGDCol_Broker
    eGDCol_AccountType
    eGDCol_NumCols
End Enum

Private Type mPrivate
    lAccountID As Long
    lAccountIdToEdit As Long
    bOK As Boolean
End Type
Private m As mPrivate

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Allows an outside caller to show the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(Optional ByVal bPassThruIfOneAccount As Boolean = False)
On Error GoTo ErrSection:

    Dim strAccountFile As String
    Dim astrAccounts As New cGdArray
    Dim lIndex As Long
    Dim rs As Recordset
    Dim strAccount As String
    
    ' Don't allow the user to delete the account if not Pro version
    If Not HasGold(False) Then Disable cmdDelete
    
    m.lAccountID = 0&
    
    ' DAJ 08/22/2007: This is obsolete code from the very first time that we were working
    ' with PATS and is now unwanted code...
    If 0 Then
        ' Add the accounts from the demo file if they do not exist already...
        strAccountFile = AddSlash(App.Path) & "Pats\Demo\TestAcct.TXT"
        If FileExist(strAccountFile) Then
            astrAccounts.FromFile strAccountFile
            For lIndex = 0 To astrAccounts.Size - 1
                strAccount = Parse(astrAccounts(lIndex), ",", 1)
                Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                        "WHERE [AccountNumber]='" & strAccount & "';", dbOpenDynaset)
                If (rs.BOF And rs.EOF) Then
                    rs.AddNew
                    rs!AccountNumber = strAccount
                    rs!Name = strAccount
                    rs!UserName = "DEMO1"
                    rs!Password = ""
                    rs!AccountType = eTT_AccountType_PATS
                    rs!StartingBalance = 0#
                    rs!CurrentBalance = 0#
                    rs!StartingDate = Date
                    rs!Broker = ""
                    rs!Comms = 0#
                    rs.Update
                End If
            Next lIndex
        End If
    End If

    ' Initialize and Load the grid
    fgAccounts.Redraw = flexRDNone
    InitGrid
    LoadGrid
    fgAccounts.Redraw = flexRDBuffered
    
    ShowButtons False
    
    ' If there are no accounts, try to get the user to enter a new account
    If VisibleRows = 0& Then NewAccount
    
    If bPassThruIfOneAccount Then
        ' Go right to the account if only one account
        If VisibleRows = 1& Then
            View
            GoTo ErrExit
        End If
    End If
    
    EnableControls
    ShowForm Me, eForm_Modal, , , ALT_GRID_ROW_COLOR
        
ErrExit:
    If m.bOK = True Then
        If FormIsLoaded("frmTTSummaryCfg") Then Unload frmTTSummaryCfg
        frmTTPositions.ShowMe m.lAccountIdToEdit
    End If
    Set rs = Nothing
    Set astrAccounts = Nothing
    Unload Me
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    Set astrAccounts = Nothing
    Unload Me
    RaiseError "frmTTAccounts.ShowMe"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowFromTrades
'' Description: Show the form from the Trade Console
'' Inputs:      Account ID open
'' Returns:     Account ID chosen
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowFromTrades(ByVal lAccountID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function

    m.lAccountID = lAccountID
    lReturn = -1&

    ' Don't allow the user to delete the account if not Pro version
    If Not HasGold(False) Then Disable cmdDelete

    fgAccounts.Redraw = flexRDNone
    InitGrid
    LoadGrid lAccountID
    fgAccounts.Redraw = flexRDBuffered
    
    ShowButtons True
    
    EnableControls
    ShowForm Me, True, , , ALT_GRID_ROW_COLOR
    
    If m.bOK = True Then
        'lReturn = CLng(ValOfText(fgAccounts.TextMatrix(fgAccounts.Row, GDCol(eGDCol_AccountID))))
        lReturn = m.lAccountIdToEdit
    End If
    
    ShowFromTrades = lReturn

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTTAccounts.ShowFromTrades", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow as outside caller to print the grid information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe()
On Error GoTo ErrSection

    PrintMe = frmPrintPreview.ShowMe("CNV TTAccounts", Me, , , , 0.75, 0.75)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTAccounts.PrintMe"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the print preview form for this grid
'' Inputs:      Arguments passed in from PrintMe
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:
    
    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim strText As String               ' Text from the current grid cell
    
    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .TextAlign = taCenterMiddle
        .Text = Me.Caption
        .TextAlign = taLeftMiddle
        .Font.Bold = False
        
        .Paragraph = ""
        .Paragraph = ""
        
        If frmPrintPreview.GoingToFile Then
            With fgAccounts
                For lRow = 0 To .Rows - 1
                    strText = ""
                    For lCol = 0 To .Cols - 1
                        If Not .ColHidden(lCol) Then
                            strText = strText & .Cell(flexcpTextDisplay, lRow, lCol) & vbTab
                        End If
                    Next lCol
                    strText = Left(strText, Len(strText) - 1) ' strip the trailing tab
                    frmPrintPreview.vp.Text = strText & vbCrLf
                Next lRow
            End With
        Else
            .RenderControl = fgAccounts.hWnd
        End If
        
        .EndDoc
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.GenerateReport"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDelete_Click
'' Description: Allow the user to delete an existing account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    Delete

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.cmdDelete_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: Allow the user to edit an existing account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    Edit

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.cmdEdit_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdExit_Click
'' Description: If the user clicks on the exit button, unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdExit_Click()
On Error GoTo ErrExit:

    m.bOK = False
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.cmdExit_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNew_Click
'' Description: Allow the user to create a new account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNew_Click()
On Error GoTo ErrSection:

    NewAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.cmdNew_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPrint_Click
'' Description: Allow the user to print out the grid
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
    RaiseError "frmTTAccounts.cmdPrint_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdView_Click
'' Description: If the user clicks on the View button, bring up the Positions
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdView_Click()
On Error GoTo ErrSection:

    View

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.cmdView_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccounts_AfterRowColChange
'' Description: Enable/Disable controls as appropriate as user changes cells
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccounts_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.fgAccounts_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccounts_DblClick
'' Description: If the user double clicks on an account, bring up the account
''              editor for that account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccounts_DblClick()
On Error GoTo ErrSection:

    With fgAccounts
        If .MouseRow >= .FixedRows Then
            .Row = .MouseRow
            .RowSel = .Row
            View
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.fgAccounts_DblClick"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccounts_KeyDown
'' Description: Allow the user to add or delete an account with the Insert and
''              Delete keys
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccounts_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        Delete
    ElseIf KeyCode = vbKeyInsert Then
        NewAccount
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTPositions.fgAccounts_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccounts_KeyPress
'' Description: If the user hits Enter on an account, bring up the account
''              editor for that account
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccounts_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then View

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.fgAccounts_KeyPress"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccounts_MouseDown
'' Description: Allow the user to get the popup menu with a right click
'' Inputs:      Mouse Button pressed, Shift/Ctrl/Alt status, Location of Click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccounts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    Dim lMouseCol As Long               ' Current mouse column in the grid
    
    With fgAccounts
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If Button = vbRightButton Then
            If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
                .RowSel = lMouseRow
                If Not .IsSelected(lMouseRow) Then .Row = lMouseRow
            End If
            
            mnuDeleteAccount.Enabled = (lMouseRow >= .FixedRows And lMouseRow < .Rows)
            mnuEditAccount.Enabled = (lMouseRow >= .FixedRows And lMouseRow < .Rows)
            mnuPrintAccounts.Enabled = .Rows > .FixedRows
            mnuViewPositions.Enabled = (lMouseRow >= .FixedRows And lMouseRow < .Rows)
            
            PopupMenu mnuPopUp
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.fgAccounts_MouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAccounts_MouseMove
'' Description: Show the user a tool-tip if hovering over fixed row
'' Inputs:      Mouse Button pressed, Shift/Ctrl/Alt status, Location of Mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAccounts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    Dim lMouseCol As Long               ' Current mouse column in the grid
    
    With fgAccounts
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow < .FixedRows And lMouseRow >= 0 Then
            .ToolTipText = SORT_BY_PREFIX & .TextMatrix(lMouseRow, lMouseCol)
        Else
            .ToolTipText = ""
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.fgAccounts_MouseMove"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Allow the user to view help if they press F1 on the form
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.Form_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, do some initialization and center it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strFont As String               ' Font information from the ini file
    Dim strPlacement As String          ' Placement info from the ini file

    Caption = "Trading Accounts"
    Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)
    
    g.Styler.StyleForm Me
    
    strPlacement = GetIniFileProperty("TTAccounts", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        Move Left, Top, 8910, 1000 '(min height will takeover)
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement, "LHTW"
    End If
    
    mnuPopUp.Visible = False
    strFont = GetIniFileProperty("TTAcounts", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString fgAccounts.Font, strFont

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.Form_Load"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hits the X, unload the form
'' Inputs:      Whether or not to Cancel the Unload, Mode of the Unload
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
    RaiseError "frmTTAccounts.Form_QueryUnload"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: If the user resizes the form, resize the controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    If LimitFormSize(Me, fraButtons.Width * 4, fraButtons.Height + fraButtons.Top * 2) Then
        Exit Sub
    End If
    
    With fgAccounts
        .Move .Left, .Top, Me.ScaleWidth - fraButtons.Width - (.Left * 3), _
            Me.ScaleHeight - (.Top * 2)
    End With
    
    With fraButtons
        .Move fgAccounts.Width + (fgAccounts.Left * 2)
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save some settings to the ini file upon unloading
'' Inputs:      Whether or not to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "TTAccounts", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "TTAccounts", FontToString(fgAccounts.Font), "Fonts", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Allow the user to change fonts on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgAccounts

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.mnuChangeFont_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuDeleteAccount_Click
'' Description: Allow the user to delete an account from the popup menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuDeleteAccount_Click()
On Error GoTo ErrSection:

    Delete

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.mnuDeleteAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditAccount_Click
'' Description: Allow the user to edit an account from the popup menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEditAccount_Click()
On Error GoTo ErrSection:

    Edit

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.mnuEditAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuNewAccount_Click
'' Description: Allow the user to create an account from the popup menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuNewAccount_Click()
On Error GoTo ErrSection:

    NewAccount

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.mnuNewAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPrintAccount_Click
'' Description: Allow the user to print accounts from the popup menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPrintAccounts_Click()
On Error GoTo ErrSection:

    PrintMe
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.mnuPrintAccounts_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuViewPositions_Click
'' Description: Allow the user to view positions from the popup menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuViewPositions_Click()
On Error GoTo ErrSection:

    View

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.mnuViewPositions_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the Accounts grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With fgAccounts
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedRows = 1
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_AccountID)) = "AccountID"
        .TextMatrix(0, GDCol(eGDCol_Number)) = "Account Number"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_StartBalance)) = "Starting Balance"
        .TextMatrix(0, GDCol(eGDCol_StartDate)) = "Start Date"
        .TextMatrix(0, GDCol(eGDCol_Broker)) = "Broker"
        .TextMatrix(0, GDCol(eGDCol_AccountType)) = "Account Type"
        
        .ColHidden(GDCol(eGDCol_AccountID)) = True
        
        .ColDataType(GDCol(eGDCol_StartDate)) = flexDTDate
        .ColFormat(GDCol(eGDCol_StartDate)) = DateFormat("Format")
        
        .ColAlignment(GDCol(eGDCol_StartDate)) = flexAlignCenterTop
        .ColAlignment(GDCol(eGDCol_Number)) = flexAlignLeftTop
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.InitGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Loads the grid with the account information from the database
'' Inputs:      Hide account ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid(Optional ByVal lHideAccountID As Long = -1&)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lRow As Long
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccounts];", dbOpenDynaset)
    With fgAccounts
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        Do While Not rs.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, GDCol(eGDCol_AccountID)) = rs!AccountID
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Number)) = rs!AccountNumber
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = rs!Name
            .TextMatrix(.Rows - 1, GDCol(eGDCol_StartBalance)) = Format(rs!StartingBalance, "$#,##0.00")
            .TextMatrix(.Rows - 1, GDCol(eGDCol_StartDate)) = rs!StartingDate
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Broker)) = NullChk(rs!Broker)
            .TextMatrix(.Rows - 1, GDCol(eGDCol_AccountType)) = g.Broker.BrokerName(rs!AccountType)
            
            .RowHidden(.Rows - 1) = (g.Broker.HideAccount(rs!AccountID) Or (rs!AccountID = lHideAccountID))
            
            rs.MoveNext
        Loop
        
        lRow = FirstVisibleRow
        If lRow > -1& Then
            .Select .FixedRows, GDCol(eGDCol_Name), .Rows - 1, GDCol(eGDCol_Name)
            .Sort = flexSortGenericAscending
            
            .Row = FirstVisibleRow
            .RowSel = .Row
        End If
        
        SetBackColors fgAccounts
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "frmTTAccounts.LoadGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim lAccountID As Long
    Dim bHasRows As Boolean
    
    If ValidRow Then
        lAccountID = CLng(ValOfText(fgAccounts.TextMatrix(fgAccounts.Row, GDCol(eGDCol_AccountID))))
    End If
    bHasRows = (VisibleRows >= 1&)

    Enable cmdEdit, bHasRows ' And (lAccountID <> m.lAccountID)
    Enable cmdDelete, bHasRows And (lAccountID <> m.lAccountID) And HasGold(False)
    Enable cmdView, bHasRows

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.EnableControls"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidRow
'' Description: Is the current row a valid row for selection?
'' Inputs:      None
'' Returns:     True if the row is valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidRow() As Boolean
On Error GoTo ErrSection:

    With fgAccounts
        If .Row >= .FixedRows And .Row < .Rows Then ValidRow = True
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTAccounts.ValidRow"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    View
'' Description: Allow the user to View the positions for this account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub View()
On Error GoTo ErrSection:

    If ValidRow Then
        m.lAccountIdToEdit = CLng(ValOfText(fgAccounts.TextMatrix(fgAccounts.Row, GDCol(eGDCol_AccountID))))
        m.bOK = True
        Hide
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.View"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewAccount
'' Description: Allow the user to create a new account if they have the PRO
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewAccount()
On Error GoTo ErrSection:

    If fgAccounts.Rows > fgAccounts.FixedRows Then
        If Not HasGold(True) Then Exit Sub
    End If
    
    m.lAccountIdToEdit = 0&
    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.NewAccount"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Edit
'' Description: Allow the user to edit an account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Edit()
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID to edit

    If ValidRow Then
        m.lAccountIdToEdit = CLng(ValOfText(fgAccounts.TextMatrix(fgAccounts.Row, GDCol(eGDCol_AccountID))))
        m.bOK = True
        Hide
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.Edit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Delete
'' Description: Allow the user to delete an account
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Delete()
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account ID to detete
    Dim Account As cPtAccount           ' Account object to delete
    
    If ValidRow Then
        lAccountID = fgAccounts.TextMatrix(fgAccounts.Row, GDCol(eGDCol_AccountID))
        
        Set Account = New cPtAccount
        If Account.Load(lAccountID) Then
            If g.Broker.DeleteAccount(Account) = True Then
                fgAccounts.RemoveItem fgAccounts.Row
                EnableControls
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.Delete"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VisibleRows
'' Description: Determine the number of visible rows in the grid
'' Inputs:      None
'' Returns:     Number of visible non-header rows in the grid
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function VisibleRows() As Long
On Error GoTo ErrSection:

    Dim lRows As Long                   ' Count of visible non-header rows
    Dim lIndex As Long                  ' Index into a for loop
    
    With fgAccounts
        lRows = 0&
        For lIndex = .FixedRows To .Rows - 1
            If .RowHidden(lIndex) = False Then lRows = lRows + 1
        Next lIndex
    End With
    
    VisibleRows = lRows

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTTAccounts.VisibleRows"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FirstVisibleRow
'' Description: Determine the first visible non-header row in the grid
'' Inputs:      None
'' Returns:     Index of the first visible non-header row or -1 if none
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FirstVisibleRow() As Long
On Error GoTo ErrSection:

    Dim lRow As Long                    ' First visible non-header row in the grid
    Dim lIndex As Long                  ' Index into a for loop
    
    With fgAccounts
        lRow = -1&
        For lIndex = .FixedRows To .Rows - 1
            If .RowHidden(lIndex) = False Then
                lRow = lIndex
                Exit For
            End If
        Next lIndex
    End With
    
    FirstVisibleRow = lRow

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "ftmTTAccounts.FirstVisibleRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowButtons
'' Description: Show or hide the buttons on the side based on the mode
'' Inputs:      From Trade Tracker?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowButtons(ByVal bFromTradeTracker As Boolean)
On Error GoTo ErrSection:

    If bFromTradeTracker Then
        cmdView.Visible = True
        cmdView.Caption = "&OK"
        cmdView.Default = True
        cmdNew.Visible = True
        cmdNew.Top = 1260
        mnuNewAccount.Visible = True
        cmdEdit.Visible = False
        mnuEditAccount.Visible = False
        cmdDelete.Visible = False
        mnuDeleteAccount.Visible = False
        cmdPrint.Visible = False
        mnuPrintAccounts.Visible = False
        cmdExit.Visible = True
        cmdExit.Caption = "&Cancel"
        cmdExit.Cancel = True
        cmdExit.Top = 420
    Else
        cmdView.Visible = True
        cmdView.Caption = "&View History"
        cmdView.Default = True
        cmdNew.Visible = True
        cmdNew.Top = 420
        mnuNewAccount.Visible = True
        cmdEdit.Visible = True
        mnuEditAccount.Visible = True
        cmdDelete.Visible = True
        mnuDeleteAccount.Visible = True
        cmdPrint.Visible = True
        mnuPrintAccounts.Visible = True
        cmdExit.Visible = True
        cmdExit.Caption = "E&xit"
        cmdExit.Cancel = True
        cmdExit.Top = 2400
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTTAccounts.ShowButtons"
    
End Sub

