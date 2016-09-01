VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmAlertsSetup 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   2115
      Left            =   3600
      TabIndex        =   0
      Top             =   180
      Width           =   1455
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
      Caption         =   "frmAlertsSetup.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAlertsSetup.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAlertsSetup.frx":0040
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdHistory 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   1680
         Width           =   1455
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
         Caption         =   "frmAlertsSetup.frx":005C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlertsSetup.frx":009A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlertsSetup.frx":00BA
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdExit 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   1260
         Width           =   1455
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
         Caption         =   "frmAlertsSetup.frx":00D6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlertsSetup.frx":0100
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlertsSetup.frx":0120
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRemoveAlert 
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   840
         Width           =   1455
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
         Caption         =   "frmAlertsSetup.frx":013C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlertsSetup.frx":0176
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlertsSetup.frx":0196
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEditAlert 
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   420
         Width           =   1455
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
         Caption         =   "frmAlertsSetup.frx":01B2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlertsSetup.frx":01E8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlertsSetup.frx":0208
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAddAlert 
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1455
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
         Caption         =   "frmAlertsSetup.frx":0224
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAlertsSetup.frx":0258
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAlertsSetup.frx":0278
         RightToLeft     =   0   'False
      End
   End
   Begin VB.Timer tmrRealTime 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   180
      Top             =   5160
   End
   Begin VSFlex7LCtl.VSFlexGrid fgAlerts 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   2895
      _cx             =   5106
      _cy             =   2566
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
      Begin VB.Menu mnuAddAlert 
         Caption         =   "Add Alert"
         Begin VB.Menu mnuAddQbAlert 
            Caption         =   "Quote Board Alert"
         End
         Begin VB.Menu mnuAddAtAlert 
            Caption         =   "Order Alert"
         End
         Begin VB.Menu mnuStatusAlert 
            Caption         =   "Status Alert"
         End
         Begin VB.Menu mnuPriceAlert 
            Caption         =   "Price Alert"
         End
         Begin VB.Menu mnuTimeAlert 
            Caption         =   "Time Alert"
         End
         Begin VB.Menu mnuChartCondition 
            Caption         =   "Chart Condition Alert"
         End
         Begin VB.Menu mnuTradeSenseAlert 
            Caption         =   "Trade Sense Alert"
         End
      End
      Begin VB.Menu mnuEditAlert 
         Caption         =   "Edit Alert"
      End
      Begin VB.Menu mnuRemoveAlert 
         Caption         =   "Remove Alert"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmAlertsSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmAlertsSetup.frm
'' Description: Allow the user to manage their alerts
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/13/2004   DAJ         Created
'' 03/09/2011   DAJ         Don't do a CheckAlert on a broker status alert
'' 02/14/2012   DAJ         New status alerts for position mismatch / auto trade disabled
'' 06/24/2013   DAJ         Timer Logging
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols            'alerts grid columns
    eGDCol_Active = 0
    eGDCol_Type
    eGDCol_Symbol
    eGDCol_Price
    eGDCol_Alert
    eGDCol_Action
    eGDCol_Index
    eGDCol_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean
    bLoadGrid As Boolean        ' Initial grid loading is in progress
    bUnloading As Boolean       ' Is the form unloading?
    
    iSortCol As Long            'to maintain sort order - issue 3993
    iSortOrder As Long
    
    BarsColl As cGdTree         ' Collection of bars for updating
End Type
Private m As mPrivate

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAddAlert_Click
'' Description: Allow the user to add an alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddAlert_Click()
On Error GoTo ErrSection:

    If ExtremeCharts = 1 Then
        mnuAddAtAlert.Visible = False
        mnuStatusAlert.Visible = False
        mnuTimeAlert.Visible = False
        mnuChartCondition.Visible = False
        mnuTradeSenseAlert.Visible = False
    End If
    
    EnableControls
    PopupMenu mnuAddAlert, , fraButtons.Left + cmdAddAlert.Left, fraButtons.Top + cmdAddAlert.Top + cmdAddAlert.Height

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.cmdAddAlert_Click"
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditAlert_Click
'' Description: Allow the user to edit an alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditAlert_Click()
On Error GoTo ErrSection:

    EditAlert fgAlerts.RowSel

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.cmdEditAlert_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdExit
'' Description: Hide the form and let ShowMe unload it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdExit_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.cmdExit_Click"

End Sub

Private Sub cmdHistory_Click()
On Error Resume Next

    PlaySoundFile           'stop sound
    
    If FormIsLoaded("frmAlertMessages") Then
        frmAlertMessages.SetFocus
    Else
        frmAlertMessages.ShowMe
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemoveAlert_Click
'' Description: Allow the user to remove an alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemoveAlert_Click()
On Error GoTo ErrSection:

    RemoveAlert

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.cmdRemoveAlert_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAlerts_AfterEdit
'' Description: After a user has changed the Active flag on the grid, save it
''              with the alert
'' Inputs:      Row and Column that the user edited
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAlerts_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index of the alert in the collection
    Dim Alert As cAlert
    
'original code: leave awhile then remove 04-02-2007
'    lIndex = fgAlerts.TextMatrix(Row, GDCol(eGDCol_Index))
'    g.Alerts(lIndex).Active = CheckedCell(fgAlerts, Row, Col)
'    If g.Alerts(lIndex).IsOrderChangeStatusAlert = False Then g.Alerts(lIndex).CheckAlert
'    If (g.Alerts(lIndex).AlertType = eGDAlertType_QuoteBoard) Then frmQuotes.DisplayAlert g.Alerts(lIndex)
           
    lIndex = fgAlerts.TextMatrix(Row, GDCol(eGDCol_Index))
    Set Alert = g.Alerts(lIndex)
    
    If Not Alert Is Nothing Then
        Alert.Active = CheckedCell(fgAlerts, Row, Col)
        If (Alert.AlertType <> eGDAlertType_Annot) And (Alert.AlertType <> eGDAlertType_Chart) And (Alert.CheckFromCheckAlerts = True) Then
            Alert.CheckAlert
        End If
        If Alert.AlertType = eGDAlertType_QuoteBoard Then frmQuotes.DisplayAlert Alert
    
        ' Need to do this so will get updated if went inactive again
        CheckedCell(fgAlerts, Row, Col) = Alert.Active
    
        'need to update the chart objects
        If Alert.AlertType = eGDAlertType_Price Then
            Alert.UpdateChartObject False
        ElseIf Not Alert.Annotation Is Nothing Then
            Alert.Annotation.UpdateAlert 2
        ElseIf Not Alert.Indicator Is Nothing Then
            Alert.Indicator.UpdateAlert 2
        End If
    End If
    
    Set Alert = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.fgAlerts_AfterEdit"
    
End Sub

Private Sub fgAlerts_AfterSort(ByVal Col As Long, Order As Integer)
On Error Resume Next

    m.iSortCol = Col
    m.iSortOrder = Order

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAlerts_BeforeEdit
'' Description: Only allow the user to edit the first column of the alerts grid
'' Inputs:      Row and Column user is trying to edit, Whether or not to cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAlerts_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> 0 Then Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.fgAlerts_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAlerts_Click
'' Description: If the user clicks on the Alerts grid, make sure that the
''              current row becomes the mouse row
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAlerts_Click()
On Error GoTo ErrSection:

    With fgAlerts
        If .MouseRow <> -1 Then
            .Row = .MouseRow
            .RowSel = .MouseRow
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.fgAlerts_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAlerts_DblClick
'' Description: When the user double clicks on an item in the Alerts grid,
''              bring up the Alert edit dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAlerts_DblClick()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Current mouse row
    
    lRow = fgAlerts.MouseRow
    If lRow >= fgAlerts.FixedRows And lRow < fgAlerts.Rows Then
        EditAlert lRow
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.fgAlerts_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAlerts_KeyDown
'' Description: If the user presses Enter on one of the items in the Alerts
''              grid, bring up the Alert edit dialog.  If the user presses
''              Delete on one of the items, remove the current row from the grid
'' Inputs:      KeyCode of the key pressed, Shift status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAlerts_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:
    
    If KeyCode = vbKeyReturn Then
        If fgAlerts.Row >= fgAlerts.FixedRows And fgAlerts.Row < fgAlerts.Rows Then
            EditAlert fgAlerts.Row
        End If
    ElseIf KeyCode = vbKeyDelete Then
        RemoveAlert
    ElseIf KeyCode = vbKeyInsert Then
        EditAlert -1&
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.fgAlerts_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAlerts_MouseDown
'' Description: If the user right clicks in the Alerts grid, show the Alerts
''              popup menu
'' Inputs:      Which Button was pressed, Shift/Ctrl/Alt status, Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAlerts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    With fgAlerts
        If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
            .Row = .MouseRow
            .RowSel = .Row
        End If
    End With

    If Button = vbRightButton Then
        mnuAddAlert.Visible = True
        EnableControls

        PopupMenu mnuPopUp
    Else
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.fgAlerts_MouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAlerts_MouseMove
'' Description: If the user moves the mouse over the Alerts grid, show a tool
''              tip with the alert text
'' Inputs:      Mouse Button pressed, Shift status, X coordinate, Y coordinate
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAlerts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GridTooltip fgAlerts

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form and its controls upon startup
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strFont As String               ' Font from the ini file
    Dim strPlacement As String          ' Placement from the ini file

    g.Styler.StyleForm Me
    
    m.bUnloading = False
    tmrRealTime.Enabled = False
    
    strFont = GetIniFileProperty("Alerts", "", "Fonts", g.strIniFile)
    If Len(strFont) = 0 Then FontFromString fgAlerts.Font, strFont
    
    strPlacement = GetIniFileProperty("AlertsSetup", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement, "LHTW"
    End If
        
    mnuPopUp.Visible = False
    Caption = "Alerts Setup"
    Me.Icon = Picture16(ToolbarIcon("ID_Alerts"), , True)
    
    Set m.BarsColl = New cGdTree
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the alerts grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw

    With fgAlerts
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .WordWrap = True
        .AllowUserResizing = flexResizeColumns
        
        .Cols = GDCol(eGDCol_NumCols)
        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        .FrozenCols = 1
        .RowHeightMax = .RowHeight(0) * 4
        
        .TextMatrix(0, GDCol(eGDCol_Active)) = "Active"
        .TextMatrix(0, GDCol(eGDCol_Type)) = "Type"
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_Price)) = "Price"
        .TextMatrix(0, GDCol(eGDCol_Alert)) = "Alerts"
        .TextMatrix(0, GDCol(eGDCol_Action)) = "Actions"
        
        .ColDataType(GDCol(eGDCol_Active)) = flexDTBoolean
        
        .ColAlignment(GDCol(eGDCol_Price)) = flexAlignRightCenter
        
        .ColHidden(GDCol(eGDCol_Index)) = True
                
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAlertsSetup.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the alerts grid from the QuoteList.ALR file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadGrid(Optional ByVal strSymbol$ = "")
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim Alert As cAlert                 ' Temporary alert object
    
    Dim Bars As New cGdBars             ' Temporary bars object
    Dim nLastGoodBar As Long            'for chart bars (eg indicator, annotation alerts)
    
    Dim iTop As Long
    Dim iBottom As Long
    Dim iRow As Long
    
    Dim iSelRow As Long
    Dim bCheckSym As Boolean
    
    Dim iPos As Long
    Dim strText As String
            
    m.bLoadGrid = True

    iTop = -1
    iBottom = -1
    iRow = -1
    iSelRow = -1
    
    If Len(strSymbol) > 0 Then
        'strip off current contract info from continuous contract symbol
        If InStr(strSymbol, "-0") <> 0 Then strSymbol = Parse(strSymbol, " ", 1)
        bCheckSym = True
    End If
    
    With fgAlerts
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        If .Rows > 0 Then
            iTop = .TopRow
            iBottom = .BottomRow
            iRow = .Row
        End If
        
        .Rows = .FixedRows
        For lIndex = 1 To g.Alerts.Count
            Set Alert = g.Alerts(lIndex)
            If Not Alert Is Nothing Then
                .Rows = .Rows + 1
                strText = Alert.EnglishString
                iPos = InStr(UCase(strText), "THEN")
                CheckedCell(fgAlerts, .Rows - 1, GDCol(eGDCol_Active)) = Alert.Active
                If iPos > 0 Then
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Alert)) = Left(strText, iPos - 3)
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Action)) = Right(strText, Len(strText) - (iPos + 4))
                Else
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Alert)) = strText
                End If
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Index)) = Str(lIndex)
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Type)) = Alert.AlertTypeText
                If Alert.AlertType = eGDAlertType_QuoteBoard And Len(Alert.TabName) > 0 Then
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = Alert.TabName
                Else
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = Alert.Symbol
                    If bCheckSym Then
                        If iSelRow < .FixedRows Then
                            If strSymbol = Alert.Symbol Then iSelRow = .Rows - 1
                        End If
                    End If
                    nLastGoodBar = 0        'reset
                    Set Bars = Alert.MyBars(nLastGoodBar)
                    If Not Bars Is Nothing Then
                        If nLastGoodBar <= 0 Then nLastGoodBar = Bars.Size - 1
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Price)) = Bars.PriceDisplay(Bars(eBARS_Close, nLastGoodBar))
                        If Len(Alert.Symbol) = 0 Or Alert.Symbol <> Bars.Prop(eBARS_Symbol) Then
                            Alert.Symbol = Bars.Prop(eBARS_Symbol)
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = Alert.Symbol
                        End If
                        If Bars.Size = 0 Then
                            lIndex = lIndex
                        End If
                    End If
                End If
            End If
        Next lIndex
        
        If iSelRow >= .FixedRows And iSelRow < .Rows Then
            .Row = iSelRow
            .Col = 2
            .Select iSelRow, 2
            .Sort = flexSortGenericAscending
            m.iSortCol = 2
            m.iSortOrder = flexSortGenericAscending
            For lIndex = .TopRow To .BottomRow
                If .TextMatrix(lIndex, 2) = strSymbol Then
                    .Select lIndex, 2
                    Exit For
                End If
            Next
            
            If lIndex >= .BottomRow And .TextMatrix(lIndex, 2) <> strSymbol Then
                .Sort = flexSortGenericDescending
                m.iSortOrder = flexSortGenericDescending
                For lIndex = .TopRow To .BottomRow
                    If .TextMatrix(lIndex, 2) = strSymbol Then
                        .Select lIndex, 2
                        Exit For
                    End If
                Next
            End If
        Else
            'restore sort order
            If m.iSortCol >= 0 And m.iSortCol < .Cols Then
                If m.iSortOrder = 1 Or m.iSortOrder = 2 Then
                    .Col = m.iSortCol
                    .Sort = m.iSortOrder
                End If
            End If
            
            If iTop > .FixedRows And iTop < .Rows Then .TopRow = iTop
            If iRow > .TopRow And iRow < .BottomRow Then .Row = iRow
        End If

        .Redraw = lRedraw
    End With
        
    EnableControls
    
    m.bLoadGrid = False
    
    FormResize Me
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the 'X', hide and let ShowMe unload
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    'If UnloadMode <> vbFormCode Then
    '    Me.Hide
    '    Cancel = True
    '    m.bOK = False
    'End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.Form_QueryUnload"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize/move the controls on the form as the form gets resized
'' Inputs:      None
'' Returns:     None
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lMinWidth As Long               ' Minimum width to allow
    Dim lMinHeight As Long              ' Minimum height to allow
    Dim lRedraw As Long

    lMinWidth = fraButtons.Width * 4
    lMinHeight = (fraButtons.Height * 2) + 180
    If LimitFormSize(Me, lMinWidth, lMinHeight) Then Exit Sub

    With fraButtons
        .Move ScaleWidth - fraButtons.Width - 60, 60
    End With
        
    With fgAlerts
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .Move 60, 60, ScaleWidth - fraButtons.Width - 180, ScaleHeight - 80
        If Not m.bLoadGrid Then
            .RowHeight(.FixedRows) = 215        'don't allow this row to "wrap"
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize eGDCol_Active, eGDCol_Price
            
            lMinWidth = .Width - (.ColWidth(eGDCol_Active) + .ColWidth(eGDCol_Type) + .ColWidth(eGDCol_Symbol) + .ColWidth(eGDCol_Price))
            lMinWidth = lMinWidth / 2
            .ColWidth(eGDCol_Alert) = lMinWidth
            .ColWidth(eGDCol_Action) = lMinWidth
            .AutoSizeMode = flexAutoSizeRowHeight
            .AutoSize GDCol(eGDCol_Alert), GDCol(eGDCol_Action)
        End If
        .Redraw = lRedraw
    End With
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditAlert
'' Description: Allow the user to edit an alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditAlert(ByVal Row As Long, Optional ByVal nAlertType As eGDAlertType = eGDAlertType_QuoteBoard)
On Error GoTo ErrSection:

    Dim Alert As New cAlert             ' Alert to send to the alerts form
    Dim lIndex As Long                  ' Index of the Alert in the collection
    
    Dim Bars As New cGdBars             ' Temporary bars object
    Dim nLastGoodBar As Long            'last good data bar for chart alerts
    
    Dim iPos As Long
    Dim strText As String
    Dim bNewPriceAlert As Boolean
    
    lIndex = -1&
        
    With fgAlerts
        If Row <> -1 Then
            lIndex = CLng(.TextMatrix(Row, GDCol(eGDCol_Index)))
            Set Alert = g.Alerts(lIndex)
            nAlertType = Alert.AlertType
            If (nAlertType = eGDAlertType_QuoteBoard) Then frmQuotes.DisplayAlert g.Alerts(lIndex), True
        Else
            Set Alert = New cAlert
            If nAlertType = eGDAlertType_AutoTrade Then
                Alert.ActionString(eAA_PlaySound) = "1,"
            ElseIf nAlertType = eGDAlertType_Price Then
                bNewPriceAlert = True
            End If
        End If
        
        If frmAlerts.ShowMe(Alert, nAlertType) = True Then
            .Redraw = flexRDNone
            If Row = -1 Then
                .Rows = .Rows + 1
                Row = .Rows - 1
                
                lIndex = g.Alerts.Add(Alert)
                .TextMatrix(Row, GDCol(eGDCol_Index)) = Str(lIndex)
            Else
                g.Alerts(lIndex) = Alert
            End If
            
            .TextMatrix(Row, eGDCol_Type) = g.Alerts(lIndex).AlertTypeText
            
            strText = g.Alerts(lIndex).EnglishString
            iPos = InStr(UCase(strText), "THEN")
            If iPos > 0 Then
                .TextMatrix(Row, GDCol(eGDCol_Alert)) = Left(strText, iPos - 3)
                .TextMatrix(Row, GDCol(eGDCol_Action)) = Right(strText, Len(strText) - (iPos + 4))
            Else
                .TextMatrix(Row, eGDCol_Alert) = strText
            End If
            
            If Alert.AlertType = eGDAlertType_QuoteBoard And Len(Alert.TabName) > 0 Then
                .TextMatrix(Row, GDCol(eGDCol_Symbol)) = Alert.TabName
            Else
                nLastGoodBar = 0    'reset
                Set Bars = Alert.MyBars(nLastGoodBar)
                If Not Bars Is Nothing Then
                    If nLastGoodBar <= 0 Then nLastGoodBar = Bars.Size - 1
                    .TextMatrix(Row, GDCol(eGDCol_Price)) = Bars.PriceDisplay(Bars(eBARS_Close, nLastGoodBar))
                    .TextMatrix(Row, GDCol(eGDCol_Symbol)) = Bars.Prop(eBARS_Symbol)
                End If
            End If
            
'JM: original code commented out to for performance optimization (remove after awhile if all okay - 02-08-2007)
'            If Alert.SymbolID <> 0& Then
'                'AddBars Alert.SymbolID
'                'Set Bars = GetBars(Alert.SymbolID)
'                .TextMatrix(Row, GDCol(eGDCol_Price)) = Bars.PriceDisplay(Bars(eBARS_Close, Bars.Size - 1))
'            Else
'                AddBars Alert.Symbol
'                Set Bars = GetBars(Alert.Symbol)
'                .TextMatrix(Row, GDCol(eGDCol_Price)) = Bars.PriceDisplay(Bars(eBARS_Close, Bars.Size - 1))
'            End If
            
            .Row = Row
            .RowSel = Row
                        
            If g.Alerts(lIndex).AlertType = eGDAlertType_Time Then g.Alerts(lIndex).CalcNextTriggerTime
            If g.Alerts(lIndex).CheckFromCheckAlerts Then
                g.Alerts(lIndex).CheckAlert , , , True
            End If
            CheckedCell(fgAlerts, Row, GDCol(eGDCol_Active)) = g.Alerts(lIndex).Active
            
            .Redraw = flexRDBuffered
        End If
        If (lIndex > 0) And (nAlertType = eGDAlertType_QuoteBoard) Then frmQuotes.DisplayAlert g.Alerts(lIndex), False
    End With
    
    EnableControls
    
    FormResize Me
    
    If bNewPriceAlert Then Alert.UpdateChartObject False
    
    
ErrExit:
    Set Alert = Nothing
    Exit Sub
    
ErrSection:
    Set Alert = Nothing
    RaiseError "frmAlertsSetup.EditAlert"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveAlert
'' Description: Allow the user to remove an alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveAlert()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index of the alert in the collection
    Dim Alert As cAlert
        
    If fgAlerts.SelectedRows = 0 Then Exit Sub
    
    lIndex = fgAlerts.TextMatrix(fgAlerts.RowSel, GDCol(eGDCol_Index))
    
    Set Alert = g.Alerts(lIndex)
        
    If Not Alert Is Nothing Then
        If Alert.AlertType = eGDAlertType_Annot Then
            If Not Alert.Annotation Is Nothing Then Alert.Annotation.UpdateAlert 0
        ElseIf Alert.AlertType = eGDAlertType_Chart Then
            If Not Alert.Indicator Is Nothing Then Alert.Indicator.UpdateAlert 0
        Else
            'all other alerts
            Alert.Active = False
            If Alert.AlertType = eGDAlertType_Price Then Alert.UpdateChartObject True
            'If g.Alerts(lIndex).IsOrderChangeStatusAlert = False Then g.Alerts(lIndex).CheckAlert
            g.Alerts.Remove lIndex
        End If
    End If
    
    ' Reload the grid so that we get the collection indexes correct...
    LoadGrid
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.RemoveAlert"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form gets unloaded, save some of the settings
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    m.bUnloading = True
    tmrRealTime.Enabled = False

    SetIniFileProperty "Alerts", FontToString(fgAlerts.Font), "Fonts", g.strIniFile
    SetIniFileProperty "AlertsSetup", GetFormPlacement(Me), "Placement", g.strIniFile
    
    Set m.BarsColl = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuAddAtAlert_Click
'' Description: Allow the user to add an alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuAddAtAlert_Click()
On Error GoTo ErrSection:

    EditAlert -1&, eGDAlertType_AutoTrade

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.mnuAddAtAlert_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuAddQbAlert_Click
'' Description: Allow the user to add an alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuAddQbAlert_Click()
On Error GoTo ErrSection:

    If Not HasLevelForAlert(eGDAlertType_QuoteBoard, True) Then Exit Sub
    
    EditAlert -1&, eGDAlertType_QuoteBoard

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.mnuAddQbAlert_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Allow the user to change the font in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgAlerts, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.mnuChangeFont_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChartCondition_Click
'' Description: Allow the user to add a chart-condition (i.e. indicator) alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChartCondition_Click()
On Error GoTo ErrSection:

    Dim frm As Form
    Dim Chart As cChart

    If Not HasLevelForAlert(eGDAlertType_Chart, True) Then Exit Sub

    Set frm = ActiveChart()
    If Not frm Is Nothing Then Set Chart = frm.Chart

    If Not Chart Is Nothing Then
        frmConditionBuilder.ShowMe Chart, , eType_Alert
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.mnuChartCondition_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditAlert_Click
'' Description: Allow the user to edit an alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEditAlert_Click()
On Error GoTo ErrSection:

    EditAlert fgAlerts.RowSel

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.mnuEditAlert_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPriceAlert_Click
'' Description: Allow the user to add an alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPriceAlert_Click()
On Error GoTo ErrSection:

    EditAlert -1, eGDAlertType_Price

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.mnuPriceAlert_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuRemoveAlert_Click
'' Description: Allow the user to remove an alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRemoveAlert_Click()
On Error GoTo ErrSection:

    RemoveAlert

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.mnuRemoveAlert_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls as applicable
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim bEnable As Boolean              ' Whether to enable the control or not

    With fgAlerts
        bEnable = (.RowSel >= .FixedRows) And (.RowSel < .Rows)
    End With
        
    Enable cmdEditAlert, bEnable
    Enable mnuEditAlert, bEnable
    Enable cmdRemoveAlert, bEnable
    Enable mnuRemoveAlert, bEnable
    
    'Enable mnuAddAtAlert, (frmTTSummary.fgTradeItems.Rows > frmTTSummary.fgTradeItems.FixedRows) Or (frmTTSummary.fgOrders.Rows > frmTTSummary.fgOrders.FixedRows)
    ''mnuAddAtAlert.Visible = FileExist(AddSlash(App.Path) & "AutoTrade.FLG")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Array of Symbols, Array of Fields
'' Returns:     True
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Optional ByVal strSymbol$ = "") As Boolean
On Error GoTo ErrSection:

    InfBox "Loading Alerts.  Please wait...", , , "Loading Alerts...", True

    m.iSortCol = -1
    m.iSortOrder = -1
    
    fgAlerts.Redraw = flexRDNone
    InitGrid
    LoadGrid strSymbol
    
    fgAlerts.Redraw = flexRDBuffered
    
    tmrRealTime.Interval = frmQuotes.tmrRealTime.Interval
    tmrRealTime.Enabled = g.RealTime.Active
    
    ShowForm Me, eForm_Nonmodal, frmMain
    InfBox ""
    
ErrExit:
    ShowMe = True
    Exit Function
    
ErrSection:
    InfBox ""
    Unload Me
    RaiseError "frmAlertsSetup.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the Print Preview form to allow printing the alerts
'' Inputs:      Args
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable fgAlerts
        Else
            .RenderControl = fgAlerts.hWnd
        End If

        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.GenerateReport"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Bring up the Print Preview form with the Alerts
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection
    
    PrintMe = frmPrintPreview.ShowMe("CNV Alerts", frmAlertsSetup, , 0.75, 0.75, 0.75, 0.75, False)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAlertsSetup.PrintMe"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuStatusAlert_Click
'' Description: Allow the user to create a status alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuStatusAlert_Click()
On Error GoTo ErrSection:

    EditAlert -1&, eGDAlertType_Status

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.mnuStatusAlert_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateHistory
'' Description: Add alert history item to history grid
'' Inputs:      dDate       - date time
''              strType     - alert type
''              strAlert    - alert string
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateHistory(ByVal dDate#, ByVal strType$, ByVal strAlert$)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.UpdateHistory"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTimeAlert_Click
'' Description: Allow the user to add a time alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTimeAlert_Click()
On Error GoTo ErrSection:

    EditAlert -1, eGDAlertType_Time

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.mnuTimeAlert_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuTradeSenseAlert_Click
'' Description: Allow the user to add a tradesense alert
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuTradeSenseAlert_Click()
On Error GoTo ErrSection:

    If Not HasLevelForAlert(eGDAlertType_TradeSense, True) Then Exit Sub
    
    EditAlert -1, eGDAlertType_TradeSense

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.mnuTradeSenseAlert_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrRealTime_Timer
'' Description: Update the prices in the Bars collection with the stream
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrRealTime_Timer()
On Error GoTo ErrSection:
    
    Dim i&, lAlertIndex&
    Dim Alert As cAlert
    Dim Bars As cGdBars
    Dim nLastGoodBar As Long
        
    TimerStart "frmAlertsSetup.tmrRealTime"
    If InStr(UCase(frmStatus.Caption), "RETRIEVING DATA") <> 0 Then
        If frmStatus.IsBusy Then
            Exit Sub
        End If
    End If
    
    If (m.bLoadGrid = False) And (m.bUnloading = False) Then
        With fgAlerts
            If .TopRow >= .FixedRows And .TopRow < .Rows Then
                If .BottomRow >= .FixedRows And .BottomRow < .Rows Then
                    If .BottomRow >= .TopRow Then
                        For i = .TopRow To .BottomRow
                            lAlertIndex = CLng(ValOfText(.TextMatrix(i, GDCol(eGDCol_Index))))
                            Set Alert = g.Alerts(lAlertIndex)
                            If Not Alert Is Nothing Then
                                nLastGoodBar = 0        'reset
                                Set Bars = Alert.MyBars(nLastGoodBar)
                                If Not Bars Is Nothing Then
                                    If nLastGoodBar <= 0 Then nLastGoodBar = Bars.Size - 1
                                    If (Alert.SymbolID = Bars.Prop(eBARS_SymbolID)) Or (Alert.Symbol = Bars.Prop(eBARS_Symbol)) Then
                                        ChangeCell fgAlerts, i, GDCol(eGDCol_Price), Bars.PriceDisplay(Bars(eBARS_Close, nLastGoodBar))
                                    End If
                                End If
                            End If
                        Next
                        ClearUpdatedColors
                    End If
                End If
            End If
        End With
    End If
    TimerEnd "frmAlertsSetup.tmrRealTime", tmrRealTime.Interval
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.tmrRealTime_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadBars
'' Description: Load and splice the bars for a particular symbol or symbol ID
'' Inputs:      Bars to Load, Symbol or Symbol ID to load
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadBars(Bars As cGdBars, ByVal vSymbolOrSymbolID As Variant, ByVal strPeriod As String, Optional ByVal bAddToRT As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim Bars2 As New cGdBars            ' Temporary Bars structure
    Dim bReturn As Boolean              ' Return value
    
    Set Bars2 = frmQuotes.GetBars(vSymbolOrSymbolID, strPeriod)
    If Bars2 Is Nothing Then
        Bars.ArrayMask = eBARS_EodBidAsk
        bReturn = DM_GetBars(Bars, vSymbolOrSymbolID, strPeriod, LastDailyDownload - 5, , , False)
    ElseIf Bars2.Size > 0 Then
        Set Bars = Bars2.MakeCopy
    Else
        Bars.ArrayMask = eBARS_EodBidAsk
        bReturn = DM_GetBars(Bars, vSymbolOrSymbolID, strPeriod, LastDailyDownload - 5, , , False)
    End If
    
    If bAddToRT Then
        g.RealTime.AddTickBuffer Bars
        g.RealTime.SpliceBars Bars
    End If
    
    LoadBars = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAlertsSetup.LoadBars"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddBars
'' Description: Add a symbol/period combination to the bars collection
'' Inputs:      Symbol or Symbol ID, Period
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddBars(ByVal vSymbolOrSymbolID As Variant)
On Error GoTo ErrSection:

    Exit Sub

    Dim Bars As New cGdBars             ' Temporary Bars structure

    If m.BarsColl.Exists(BarsKey(vSymbolOrSymbolID, "Daily")) = False Then
        LoadBars Bars, vSymbolOrSymbolID, "Daily"
        m.BarsColl.Add Bars, BarsKey(vSymbolOrSymbolID, "Daily")
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.AddBars"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBars
'' Description: Get the Bars for the given symbol from the collection (add it
''              to the collection if not there)
'' Inputs:      Symbol to get Bars for
'' Returns:     Bars
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetBars(ByVal vSymbolOrSymbolID As Variant) As cGdBars
On Error GoTo ErrSection:

    If m.BarsColl.Exists(BarsKey(vSymbolOrSymbolID, "Daily")) = False Then
        AddBars vSymbolOrSymbolID
    End If
    
    Set GetBars = m.BarsColl(BarsKey(vSymbolOrSymbolID, "Daily"))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAlertsSetup.GetBars"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BarsKey
'' Description: Determine the key into the bars collection
'' Inputs:      Symbol or Symbol ID, Period
'' Returns:     Key into the collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BarsKey(ByVal vSymbolOrSymbolID As Variant, ByVal strPeriod As String) As String
On Error GoTo ErrSection:

    BarsKey = Str(vSymbolOrSymbolID) & vbTab & strPeriod

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAlertsSetup.BarsKey"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshPrice
'' Description: Refresh the prices in the grid for the given symbol/period
'' Inputs:      Bars
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshPrice(ByVal Bars As cGdBars)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lAlertIndex As Long             ' Index into the alerts collection
    Dim Alert As New cAlert             ' Temporary alert object
    
    With fgAlerts
        For lIndex = .FixedRows To .Rows - 1
            lAlertIndex = CLng(ValOfText(.TextMatrix(lIndex, GDCol(eGDCol_Index))))
            Set Alert = g.Alerts(lAlertIndex)
            If Not Alert Is Nothing Then
                If (Alert.SymbolID = Bars.Prop(eBARS_SymbolID)) Or (Alert.Symbol = Bars.Prop(eBARS_Symbol)) Then
                    If Alert.Period = Bars.Prop(eBARS_PeriodicityStr) Then
                        ChangeCell fgAlerts, lIndex, GDCol(eGDCol_Price), Bars.PriceDisplay(Bars(eBARS_Close, Bars.Size - 1))
                    End If
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.RefreshPrice"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeCell
'' Description: To change text and forecolor of grid cell
'' Inputs:      Grid to change, Row and Column to change, New Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeCell(Grid As VSFlexGrid, ByVal lRow&, ByVal lCol&, ByVal strCellText$)
On Error GoTo ErrSection:

    Dim nForeColor As Long              ' Foreground color to color the cell
    Dim dTickCount As Double            ' Current tick count
    
    With Grid
        nForeColor = frmQuotes.UnchColor
        If .TextMatrix(lRow, lCol) <> strCellText Then
            .TextMatrix(lRow, lCol) = strCellText
            If tmrRealTime.Enabled Then
                nForeColor = frmQuotes.UpdateColor
                .Cell(flexcpForeColor, lRow, 0) = frmQuotes.UpdateColor
                .Cell(flexcpData, lRow, lCol) = gdTickCount
            End If
        ElseIf tmrRealTime.Enabled Then
            dTickCount = .Cell(flexcpData, lRow, lCol)
            dTickCount = gdTickCount - dTickCount
            If dTickCount >= 0 And dTickCount <= g.nUpdatedColorDuration Then
                nForeColor = frmQuotes.UpdateColor
            End If
        End If
        
        .Cell(flexcpForeColor, lRow, lCol) = nForeColor
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAlertsSetup.ChangeCell"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearUpdatedColors
'' Description: Clear the updated colors on both grids if necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearUpdatedColors()
On Error GoTo ErrSection:

    Dim lRow As Long
    Dim lCol As Long
    Dim dTickCount As Double
    Dim iSaveRedraw As Integer
    Dim bStillColor As Boolean

    With fgAlerts
        iSaveRedraw = .Redraw
        .Redraw = flexRDNone
        lCol = GDCol(eGDCol_Price)
        
        For lRow = .TopRow To .BottomRow
            If g.bUnloading Then Exit Sub
            If .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UpdateColor Then
                bStillColor = False
                If tmrRealTime.Enabled Then
                    If .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UpdateColor Then
                        ' see if has been more than 1 second since colored
                        dTickCount = .Cell(flexcpData, lRow, lCol)
                        dTickCount = gdTickCount - dTickCount
                        If dTickCount >= 0 And dTickCount <= g.nUpdatedColorDuration Then
                            bStillColor = True
                        Else
                            .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UnchColor
                        End If
                    End If
                End If
                
                ' color symbol cell only if a cell was still colored
                If Not bStillColor Then
                    .Cell(flexcpForeColor, lRow, lCol) = frmQuotes.UnchColor
                End If
            End If
        Next
        .Redraw = iSaveRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmAlertsSetup.ClearUpdatedColors"

End Sub

