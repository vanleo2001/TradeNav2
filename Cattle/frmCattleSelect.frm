VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmCattleSelect 
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1095
      Left            =   3180
      TabIndex        =   1
      Top             =   120
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
      Caption         =   "frmCattleSelect.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCattleSelect.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCattleSelect.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   600
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
         Caption         =   "frmCattleSelect.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCattleSelect.frx":0096
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCattleSelect.frx":00B6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
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
         Caption         =   "frmCattleSelect.frx":00D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCattleSelect.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCattleSelect.frx":0118
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgList 
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
Attribute VB_Name = "frmCattleSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmCattleSelect.frm
'' Description: Form for allowing user to select a lot
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
'' 11/15/2013   DAJ         Changed way to get Turnkey icon for the form
'' 11/22/2013   DAJ         Renamed frmTurnkeySelectLot to frmTurnkeySelect
'' 11/22/2013   DAJ         Import historical fills for Turnkey
'' 12/03/2013   DAJ         Added accounts mode
'' 12/30/2013   DAJ         Fix for not being able to check boxes in account mode
'' 03/07/2014   DAJ         Moved into NavCattle.DLL
'' 03/13/2014   DAJ         Fixed display account in fill text function
'' 05/22/2014   DAJ         Renamed frmTurnkeySelect to frmCattleSelect
'' 05/22/2014   DAJ         Renamed g.Turnkey to g.Cattle
'' 05/30/2014   DAJ         Utilized new accounts object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_Use = 0
    eGDCol_ID = 1
    eGDCol_Text = 2
    eGDCol_Status = 3
    eGDCol_Key = 4
    
    eGDCol_NumCols
End Enum

Private Enum eGDCattleSelectModes
    eGDCattleSelectMode_Lot
    eGDCattleSelectMode_Fills
    eGDCattleSelectMode_Accounts
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user OK the dialog?
    
    nMode As eGDCattleSelectModes       ' Mode for the form
    strFeedYardLotID As String          ' Feed Yard Lot ID to return
    Fills As cGdTree                    ' Fills to return
End Type
Private m As mPrivate

Public Property Get FeedYardLotID() As String
    FeedYardLotID = m.strFeedYardLotID
End Property

Public Property Get Fills() As cGdTree
    Set Fills = m.Fills
End Property

Private Property Get GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeLot
'' Description: Setup and show the form to select a lot
'' Inputs:      Feed Yard Lot ID
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeLot(Optional ByVal strFeedYardLotID As String = "") As Boolean
On Error GoTo ErrSection:

    m.nMode = eGDCattleSelectMode_Lot
    Caption = "Select Lot"

    InitGrid
    LoadGridLots strFeedYardLotID
    
    m.strFeedYardLotID = ""

    ShowForm Me, eForm_Modal, g.frmMain
    
    If m.bOK = True Then
        m.strFeedYardLotID = fgList.TextMatrix(fgList.RowSel, GDCol(eGDCol_ID))
    End If
    
    ShowMeLot = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmCattleSelect.ShowMeLot"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeFills
'' Description: Setup and show the form to select fills
'' Inputs:      Fills
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeFills(ByVal Fills As cGdTree) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    m.nMode = eGDCattleSelectMode_Fills
    Caption = "Select Fills"
    Set m.Fills = New cGdTree

    InitGrid
    LoadGridFills Fills
    
    ShowForm Me, eForm_Modal, g.frmMain
    
    m.Fills.Clear
    If m.bOK Then
        With fgList
            For lIndex = .FixedRows To .Rows - 1
                If CheckedCell(fgList, lIndex, GDCol(eGDCol_Use)) = True Then
                    m.Fills.Add .RowData(lIndex)
                End If
            Next lIndex
        End With
    End If
    
    ShowMeFills = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmCattleSelect.ShowMeFills"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeAccounts
'' Description: Setup and show the form to select accounts
'' Inputs:      Broker Accounts, Associated Accounts, Feedyard
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeAccounts(ByVal BrokerAccounts As cGdTree, AssociatedAccounts As cGdTree, ByVal strFeedyard As String) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    m.nMode = eGDCattleSelectMode_Accounts
    Caption = "Accounts to use with " & strFeedyard

    InitGrid
    LoadGridAccounts BrokerAccounts, AssociatedAccounts
    
    ShowForm Me, eForm_Modal, g.frmMain
    
    If m.bOK Then
        AssociatedAccounts.Clear
        With fgList
            For lIndex = .FixedRows To .Rows - 1
                If CheckedCell(fgList, lIndex, GDCol(eGDCol_Use)) = True Then
                    AssociatedAccounts.Add .RowData(lIndex), .TextMatrix(lIndex, GDCol(eGDCol_Key))
                End If
            Next lIndex
        End With
    End If
    
    ShowMeAccounts = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmCattleSelect.ShowMeAccounts"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Handle the user clicking on the Cancel button
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
    RaiseError "frmCattleSelect.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Handle the user clicking on the OK button
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
    RaiseError "frmCattleSelect.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgList_AfterRowColChange
'' Description: Handle the user changing the current cell
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    With fgList
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleSelect.fgList_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgList_Click
'' Description: Handle a click in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgList_Click()
On Error GoTo ErrSection:

    Dim lMouseCol As Long               ' Mouse column in the grid
    Dim lMouseRow As Long               ' Mouse row in the grid
    
    If (m.nMode = eGDCattleSelectMode_Fills) Or (m.nMode = eGDCattleSelectMode_Accounts) Then
        With fgList
            lMouseCol = .MouseCol
            lMouseRow = .MouseRow
            
            If lMouseCol = GDCol(eGDCol_Use) Then
                CheckedCell(fgList, lMouseRow, lMouseCol) = Not CheckedCell(fgList, lMouseRow, lMouseCol)
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleSelect.fgList_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgList_DblClick
'' Description: Handle a double click in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgList_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Mouse row in the grid
    
    If m.nMode = eGDCattleSelectMode_Lot Then
        With fgList
            lMouseRow = .MouseRow
            
            If (lMouseRow >= .FixedRows) And (lMouseRow < .Rows) Then
                .Row = lMouseRow
                .RowSel = lMouseRow
                
                m.bOK = True
                Hide
            End If
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleSelect.fgList_DblClick"
    
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

    Icon = g.AppBridge.Picture16(g.Cattle.IconName)
    
    g.Styler.StyleForm Me
    
    PlaceForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleSelect.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Handle the user clicking on the X
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
    RaiseError "frmCattleSelect.Form_QueryUnload"
    
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

    Dim lSpace As Long                  ' Space between controls
    
    lSpace = 120

    If LimitFormSize(Me, (fraButtons.Width * 3) + (lSpace * 3), fraButtons.Height + (lSpace * 2)) = False Then
        With fraButtons
            .Move ScaleWidth - .Width - lSpace, lSpace
        End With
        
        With fgList
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
    RaiseError "frmCattleSelect.Form_Unload"
    
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

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid

    With fgList
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        SetupGrid fgList, eGridMode_List
        
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
                
        Select Case m.nMode
            Case eGDCattleSelectMode_Lot
                .Rows = 0
                .FixedRows = 0
                
                .ColHidden(GDCol(eGDCol_Use)) = True
                .ColHidden(GDCol(eGDCol_ID)) = True
                .ColHidden(GDCol(eGDCol_Status)) = True
                .ColHidden(GDCol(eGDCol_Key)) = True
                
            Case eGDCattleSelectMode_Fills
                .Rows = 0
                .FixedRows = 0
                
                .ColHidden(GDCol(eGDCol_Use)) = False
                .ColHidden(GDCol(eGDCol_ID)) = False
                .ColDataType(GDCol(eGDCol_Use)) = flexDTBoolean
                .ColHidden(GDCol(eGDCol_Status)) = True
                .ColHidden(GDCol(eGDCol_Key)) = True
            
            Case eGDCattleSelectMode_Accounts
                .Rows = 1
                .FixedRows = 1
                
                .TextMatrix(0, GDCol(eGDCol_Use)) = "Use"
                .TextMatrix(0, GDCol(eGDCol_ID)) = "Account"
                .TextMatrix(0, GDCol(eGDCol_Text)) = "Broker"
                .TextMatrix(0, GDCol(eGDCol_Status)) = "Status"
                .TextMatrix(0, GDCol(eGDCol_Key)) = "Key"
                
                .ColHidden(GDCol(eGDCol_Use)) = False
                .ColHidden(GDCol(eGDCol_ID)) = False
                .ColAlignment(GDCol(eGDCol_ID)) = flexAlignLeftTop
                .ColDataType(GDCol(eGDCol_Use)) = flexDTBoolean
                .ColHidden(GDCol(eGDCol_Status)) = False
                .ColHidden(GDCol(eGDCol_Key)) = True
                                
        End Select
    
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleSelect.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridLots
'' Description: Load the grid with feed yard lots
'' Inputs:      Feed Yard Lot ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridLots(Optional ByVal strFeedYardLotID As String = "")
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim Lots As cGdTree                 ' Collection of Lots
    Dim Lot As cBrokerMessage           ' Lot object
    Dim lSelect As Long                 ' Row to select

    With fgList
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        lSelect = .FixedRows
        
        Set Lots = g.Cattle.Lots
        For lIndex = 1 To Lots.Count
            Set Lot = Lots(lIndex)
            
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = Lot("FeedYardLotID")
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Text)) = g.Cattle.LotDisplay(Lot)
            
            If Lot("FeedYardLotID") = strFeedYardLotID Then
                lSelect = .Rows - 1
            End If
        Next lIndex
        
        .Row = lSelect
        .RowSel = lSelect
        
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleSelect.LoadGridLots"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridFills
'' Description: Load the grid with fills
'' Inputs:      Fills
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridFills(ByVal Fills As cGdTree)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cBrokerMessage          ' Fill object

    With fgList
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        For lIndex = 1 To Fills.Count
            Set Fill = Fills(lIndex)
            
            .Rows = .Rows + 1
            .RowData(.Rows - 1) = Fill
            
            .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = Fill("BrokerFillID")
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Text)) = FillText(Fill)
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False
        .Redraw = nRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleSelect.LoadGridFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillText
'' Description: Build the text for a fill
'' Inputs:      Fill
'' Returns:     Text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FillText(ByVal Fill As cBrokerMessage) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    If Fill("IsBuy") = "1" Then
        strReturn = "Bought "
    Else
        strReturn = "Sold "
    End If
    strReturn = strReturn & Str(Fill("Quantity")) & " " & Fill("Symbol") & " at " & g.AppBridge.PriceDisplay(Val(Fill("Price")), Fill("Symbol"))
    strReturn = strReturn & " on " & DateFormat(Val(Fill("FillTime")), MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
    strReturn = strReturn & " in account " & g.Cattle.Accounts.DisplayAccountNumber(Fill("BrokerAccountID"))
    
    FillText = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleSelect.FillText"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridAccounts
'' Description: Load the grid
'' Inputs:      Broker Accounts, Associated Accounts
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridAccounts(ByVal BrokerAccounts As cGdTree, ByVal AssociatedAccounts As cGdTree)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' State of the grid's redraw setting
    Dim lIndex As Long                  ' Index into a for loop
    Dim Account As cBrokerMessage       ' Broker account
    Dim strKey As String                ' Key into the collection
    Dim AccountMap As cGdTree           ' Account to Row map
    Dim lRow As Long                    ' Row in the grid
    
    With fgList
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
            CheckedCell(fgList, lRow, GDCol(eGDCol_Use)) = False
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
        
            CheckedCell(fgList, lRow, GDCol(eGDCol_Use)) = True
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleSelect.LoadGridAccounts"
    
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
    
    With fgList
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .RowData(lRow) = Account
        
        If Len(Account("FcmNumber")) > 0 Then
            .TextMatrix(lRow, GDCol(eGDCol_ID)) = Account("FcmNumber")
        Else
            .TextMatrix(lRow, GDCol(eGDCol_ID)) = Account("Number")
        End If
        .TextMatrix(lRow, GDCol(eGDCol_Text)) = g.AppBridge.BrokerName(CLng(Val(Account("Broker"))))
        .TextMatrix(lRow, GDCol(eGDCol_Status)) = g.BrokerEnums.ConnectionStatusToString(g.AppBridge.ConnectionStatusForAccount(Account("Number"), True))
        .TextMatrix(lRow, GDCol(eGDCol_Key)) = strKey
    
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleSelect.AccountToGrid"
    
End Sub

