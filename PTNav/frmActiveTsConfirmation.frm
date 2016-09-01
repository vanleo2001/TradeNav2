VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmActiveTsConfirmation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraEntryButtons 
      Height          =   495
      Left            =   1980
      TabIndex        =   7
      Top             =   2220
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
      Caption         =   "frmActiveTsConfirmation.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmActiveTsConfirmation.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmActiveTsConfirmation.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Index           =   1
         Left            =   1320
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
         Caption         =   "frmActiveTsConfirmation.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmActiveTsConfirmation.frx":0096
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsConfirmation.frx":00B6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSubmitGroups 
         Height          =   495
         Left            =   0
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
         Caption         =   "frmActiveTsConfirmation.frx":00D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmActiveTsConfirmation.frx":010E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsConfirmation.frx":012E
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgGroups 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   4395
      _cx             =   7752
      _cy             =   2143
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
   Begin HexUniControls.ctlUniFrameWL fraExitButtons 
      Height          =   495
      Left            =   180
      TabIndex        =   3
      Top             =   2220
      Width           =   3855
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
      Caption         =   "frmActiveTsConfirmation.frx":014A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmActiveTsConfirmation.frx":0176
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmActiveTsConfirmation.frx":0196
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Index           =   0
         Left            =   2640
         TabIndex        =   6
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
         Caption         =   "frmActiveTsConfirmation.frx":01B2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmActiveTsConfirmation.frx":01E0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsConfirmation.frx":0200
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancelGroups 
         Height          =   495
         Left            =   1320
         TabIndex        =   5
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
         Caption         =   "frmActiveTsConfirmation.frx":021C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmActiveTsConfirmation.frx":0258
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsConfirmation.frx":0278
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdParkGroups 
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
         Caption         =   "frmActiveTsConfirmation.frx":0294
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmActiveTsConfirmation.frx":02CC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmActiveTsConfirmation.frx":02EC
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP lblQuestion 
      Height          =   255
      Left            =   180
      Top             =   1800
      Width           =   3615
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
      Caption         =   "frmActiveTsConfirmation.frx":0308
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmActiveTsConfirmation.frx":038E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmActiveTsConfirmation.frx":03AE
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblDescription 
      Height          =   255
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
      Caption         =   "frmActiveTsConfirmation.frx":03CA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmActiveTsConfirmation.frx":0460
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmActiveTsConfirmation.frx":0480
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmActiveTsConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmActiveTsConfirmation.cls
'' Description: Have user confirm actions for active TradeSense order groups
''              upon Trade Navigator close or open
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 07/05/2010   DAJ         Finished Entry dialog and fixed close for broker
'' 10/28/2010   DAJ         Hide the form before submit/park/cancel
'' 12/14/2010   DAJ         UpdateLastModified so that Orders UI updates better
'' 02/11/2013   DAJ         Cancel left over orders when the form goes away
'' 10/27/2014   DAJ         Don't allow the Submit button if no items selected
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_Select = 0
    eGDCol_Name
    eGDCol_Symbol
    eGDCol_Account
    eGDCol_Reason
    eGDCol_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user OK the dialog or cancel it?
    strReason As String                 ' Reason for being called
    nBroker As eTT_AccountType          ' Broker for this instance of the dialog
    bExit As Boolean                    ' Are we being called in exit mode?
End Type
Private m As mPrivate

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

Private Property Get PocFile() As String
    PocFile = AddSlash(App.Path) & "Custom\Poc.Sav"
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeForExit
'' Description: Show the form in exit mode
'' Inputs:      Reason, Broker
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeForExit(ByVal strReason As String, Optional ByVal nBroker As eTT_AccountType = -1&) As Boolean
On Error GoTo ErrSection:

    m.strReason = strReason
    m.nBroker = nBroker
    
    lblDescription.Caption = "The following TradeSense order groups are currently active:"
    lblQuestion.Caption = "Would you like to Cancel or Park all of the groups?"
    fraExitButtons.Visible = True
    fraEntryButtons.Visible = False

    m.bExit = True
    InitGrid
    LoadGridForExit
    
    EnableControls
    
    If fgGroups.Rows = fgGroups.FixedRows Then
        m.bOK = True
        KillFile PocFile
    Else
        ShowForm Me, eForm_Modal, frmMain, , ALT_GRID_ROW_COLOR
    End If
    
    ShowMeForExit = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmActiveTsConfirmation.ShowMeForExit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeForEntry
'' Description: Show the form in entry mode
'' Inputs:      Broker
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeForEntry(Optional ByVal nBroker As eTT_AccountType = -1&) As Boolean
On Error GoTo ErrSection:

    m.nBroker = nBroker

    lblDescription.Caption = "The following TradeSense order groups are currently parked:"
    lblQuestion.Caption = "Would you like to Submit the selected groups?"
    fraExitButtons.Visible = False
    fraEntryButtons.Visible = True

    m.bExit = False
    InitGrid
    LoadGridForEntry
    
    EnableControls
    
    If fgGroups.Rows = fgGroups.FixedRows Then
        m.bOK = True
    Else
        ShowForm Me, eForm_Modal, frmMain
    End If
    
    ShowMeForEntry = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmActiveTsConfirmation.ShowMeForEntry"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Close the form without doing anything
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click(Index As Integer)
On Error GoTo ErrSection:

    m.bOK = False
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsConfirmation.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancelGroups_Click
'' Description: Close the form and cancel all active groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancelGroups_Click()
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout variable

    Visible = False
    g.TsoGroups.CancelAllActive m.strReason, m.nBroker
    
    lTimeOut = 0
    Do While (g.TsoGroups.HasWorkingGroups(m.nBroker)) And (lTimeOut < 30)
        Sleep 1
        lTimeOut = lTimeOut + 1&
    Loop

    m.bOK = Not g.TsoGroups.HasWorkingGroups(m.nBroker)
    If m.bOK = False Then
        InfBox "Timed out while waiting for active TradeSense order groups to cancel", "!", , "Error"
    End If
    
    g.TsoGroups.UpdateLastModified
    
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsConfirmation.cmdCancelGroups_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdParkGroups_Click
'' Description: Close the form and park all active groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdParkGroups_Click()
On Error GoTo ErrSection:

    Dim lTimeOut As Long                ' Timeout variable
    Dim strParked As String             ' List of items that were parked
 
    Visible = False
    strParked = g.TsoGroups.ParkAllActive(m.nBroker, "User chose to park from Active Confirmation")

    lTimeOut = 0
    Do While (g.TsoGroups.HasWorkingGroups(m.nBroker)) And (lTimeOut < 30)
        Sleep 1
        lTimeOut = lTimeOut + 1&
    Loop

    m.bOK = Not g.TsoGroups.HasWorkingGroups(m.nBroker)
    If m.bOK = False Then
        InfBox "Timed out while waiting for active TradeSense order groups to park", "!", , "Error"
    Else
        FileFromString PocFile, strParked
    End If
    
    g.TsoGroups.UpdateLastModified
    
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsConfirmation.cmdParkGroups_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSubmitGroups_Click
'' Description: Close the form and submit all selected groups
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSubmitGroups_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    Visible = False
    With fgGroups
        For lIndex = .FixedRows To .Rows - 1
            If CheckedCell(fgGroups, lIndex, GDCol(eGDCol_Select)) Then
                If TypeOf .RowData(lIndex) Is cActiveTsOrderGroup Then
                    g.TsoGroups.SubmitParkedGroup .RowData(lIndex)
                End If
            End If
        Next lIndex
    End With
    
    g.TsoGroups.UpdateLastModified
    
    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsConfirmation.cmdSubmitGroups_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgGroups_BeforeMouseDown
'' Description: Handle the user clicking in the grid
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, X Location of the mouse,
''              Y location of the mouse, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgGroups_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Row in the grid that the user is clicking in
    Dim lMouseCol As Long               ' Column in the grid that the user is clicking in

    If Button = vbLeftButton Then
        With fgGroups
            lMouseRow = .MouseRow
            lMouseCol = .MouseCol
            
            If (mFlexGrid.ValidGridCol(fgGroups, lMouseCol) = True) And (mFlexGrid.ValidGridRow(fgGroups, lMouseRow) = True) Then
                If m.bExit = False Then
                    If lMouseCol = GDCol(eGDCol_Select) Then
                        If Len(fgGroups.TextMatrix(lMouseRow, GDCol(eGDCol_Reason))) > 0 Then
                            InfBox "You cannot activate '" & fgGroups.TextMatrix(lMouseRow, GDCol(eGDCol_Name)) & "' because " & fgGroups.TextMatrix(lMouseRow, GDCol(eGDCol_Reason)), "!", , "Activation Error"
                        Else
                            mFlexGrid.ToggleCell fgGroups, lMouseRow, lMouseCol
                            EnableControls
                        End If
                    End If
                End If
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsConfirmation.fgGroups_BeforeMouseDown"
    
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

    Caption = "Active TradeSense Order Groups"
    Icon = Picture16(ToolbarIcon("kTradeSenseOrders"))
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsConfirmation.Form_Load"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the 'X', let ShowMe unload the form
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
    RaiseError "frmActiveTsConfirmation.Form_QueryUnload"
    
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

    With lblQuestion
        .Move (ScaleWidth / 2) - (.Width / 2)
    End With

    With fraExitButtons
        .Move (ScaleWidth / 2) - (.Width / 2)
    End With

    With fraEntryButtons
        .Move (ScaleWidth / 2) - (.Width / 2)
    End With

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
        .Editable = flexEDNone
        .ExtendLastCol = True
        
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .Rows = 1
        .FixedRows = 1
        
        .TextMatrix(0, GDCol(eGDCol_Select)) = "On"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_Account)) = "Account"
        .TextMatrix(0, GDCol(eGDCol_Reason)) = "Reason"
        
        .ColDataType(GDCol(eGDCol_Select)) = flexDTBoolean
        .ColHidden(GDCol(eGDCol_Select)) = m.bExit
        .ColHidden(GDCol(eGDCol_Reason)) = True
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsConfirmation.InitGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridForExit
'' Description: Load the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridForExit()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim grp As cActiveTsOrderGroup      ' Active TradeSense order group object

    With fgGroups
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        For lIndex = 1 To g.TsoGroups.Count
            Set grp = g.TsoGroups(lIndex)
            If grp.Submitted Then
                If IncludeGroup(grp) Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = grp.tsOrderGroup.Name
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = grp.Symbol
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Account)) = g.Broker.AccountNameForID(grp.AccountID)
                End If
            End If
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActiveTsConfirmation.LoadGridForExit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridForEntry
'' Description: Load the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridForEntry()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim grp As cActiveTsOrderGroup      ' Active TradeSense order group object
    Dim strParkedOnClose As String      ' List of items that were parked on close
    Dim bCanActivate As Boolean         ' Can this be activated?
    Dim strReason As String             ' Reason cannot activate

    strParkedOnClose = "|" & FileToString(PocFile) & "|"

    With fgGroups
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        For lIndex = 1 To g.TsoGroups.Count
            Set grp = g.TsoGroups(lIndex)
            If Not grp.Submitted Then
                If IncludeGroup(grp) Then
                    grp.CancelLeftOverOrders
                    
                    .Rows = .Rows + 1
                    
                    .RowData(.Rows - 1) = grp
                    
                    bCanActivate = grp.CanActivate(strReason)
                    
                    If bCanActivate Then
                        CheckedCell(fgGroups, .Rows - 1, GDCol(eGDCol_Select)) = (InStr(strParkedOnClose, "|" & grp.Key & "|") <> 0)
                    Else
                        CheckedCell(fgGroups, .Rows - 1, GDCol(eGDCol_Select)) = False
                    End If
                    
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = grp.tsOrderGroup.Name
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = grp.Symbol
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Account)) = g.Broker.AccountNameForID(grp.AccountID)
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Reason)) = strReason
                    
                    If bCanActivate Then
                        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = .Cell(flexcpForeColor, 0, 0)
                    Else
                        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
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
    RaiseError "frmActiveTsConfirmation.LoadGridForEntry"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IncludeGroup
'' Description: Determine whether to include the group in the grid or not
'' Inputs:      Group
'' Returns:     True if Include, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IncludeGroup(ByVal grp As cActiveTsOrderGroup) As Boolean
On Error GoTo ErrSection:

    IncludeGroup = ((m.nBroker = -1&) Or (m.nBroker = grp.Broker))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmActiveTsConfirmation.IncludeGroup"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasItemsChecked
'' Description: Determine whether the user has any items checked
'' Inputs:      None
'' Returns:     True if items checked, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HasItemsChecked() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    With fgGroups
        For lIndex = .FixedRows To .Rows - 1
            If CheckedCell(fgGroups, lIndex, GDCol(eGDCol_Select)) = True Then
                bReturn = True
                Exit For
            End If
        Next lIndex
    End With
    
    HasItemsChecked = bReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmActiveTsConfirmation.HasItemsChecked"
    
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

    If m.bExit = False Then
        Enable cmdSubmitGroups, HasItemsChecked
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmActiveTsConfirmation.EnableControls"
    
End Sub

