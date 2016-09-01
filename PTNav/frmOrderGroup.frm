VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmOrderGroup 
   Caption         =   "Form1"
   ClientHeight    =   4230
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniTextBoxXP txtDescription 
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   3375
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmOrderGroup.frx":0000
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
      MultiLine       =   -1  'True
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmOrderGroup.frx":0020
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderGroup.frx":0040
   End
   Begin HexUniControls.ctlUniTextBoxXP txtName 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmOrderGroup.frx":005C
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
      Tip             =   "frmOrderGroup.frx":007C
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderGroup.frx":009C
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   2895
      Left            =   3240
      TabIndex        =   5
      Top             =   1200
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
      Caption         =   "frmOrderGroup.frx":00B8
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOrderGroup.frx":00E4
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderGroup.frx":0104
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   0
         TabIndex        =   0
         Top             =   2400
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
         Caption         =   "frmOrderGroup.frx":0120
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderGroup.frx":014E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderGroup.frx":016E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSave 
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   1860
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
         Caption         =   "frmOrderGroup.frx":018A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderGroup.frx":01B4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderGroup.frx":01D4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   495
         Left            =   0
         TabIndex        =   8
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
         Caption         =   "frmOrderGroup.frx":01F0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderGroup.frx":021E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderGroup.frx":023E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   495
         Left            =   0
         TabIndex        =   7
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
         Caption         =   "frmOrderGroup.frx":025A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderGroup.frx":0284
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderGroup.frx":02A4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNew 
         Height          =   495
         Left            =   0
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
         Caption         =   "frmOrderGroup.frx":02C0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderGroup.frx":02E8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderGroup.frx":0308
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgOrders 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   1200
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
   Begin HexUniControls.ctlUniLabelXP lblDescription 
      Height          =   255
      Left            =   120
      Top             =   480
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
      Caption         =   "frmOrderGroup.frx":0324
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOrderGroup.frx":035E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderGroup.frx":037E
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblName 
      Height          =   255
      Left            =   120
      Top             =   135
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
      Caption         =   "frmOrderGroup.frx":039A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmOrderGroup.frx":03C6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderGroup.frx":03E6
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuNew 
         Caption         =   "Create Order"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Order"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Order"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmOrderGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmOrderGroup.frm
'' Description: Allow the user to edit/create an order group
''
'' Author:      Genesis Financial Data Services
''              425 WindChime Place
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_ID = 0
    eGDCol_OrderText
    eGDCol_DisplayText
    eGDCol_NumCols
End Enum

Private Enum eGDOrderFields
    eGDOrderField_Buy = 0
    eGDOrderField_Quantity
    eGDOrderField_Symbol
    eGDOrderField_OrderType
    eGDOrderField_StopPrice
    eGDOrderField_LimitPrice
    eGDOrderField_Account
    eGDOrderField_Expiration
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user press Save or Cancel?
    lOrderGroupID As Long               ' Order Group ID for this order group
End Type
Private m As mPrivate

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

Private Function OrderField(ByVal field As eGDOrderFields) As Long
    OrderField = field
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Set up the controls and show the form
'' Inputs:      Order Group ID
'' Returns:     Order Group ID if Save, Zero if Cancel
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal lOrderGroupID As Long) As Long
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    m.lOrderGroupID = lOrderGroupID
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderGroups] WHERE [OrderGroupID]=" & Str(lOrderGroupID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        txtName.Text = rs!Name
        txtDescription.Text = rs!Description
    End If
    
    InitGrid
    LoadGrid
    
    EnableControls
    SetEditorCaption Me, "Order Group", txtName.Text

    ' This form calls frmTTEditOrder which gets shown in an "Act Modal" state,
    ' so this must be "Act Modal" as well...
    ShowForm Me, eForm_ActModal, frmMain, , ALT_GRID_ROW_COLOR
    
    If m.bOK Then
        Save
        ShowMe = m.lOrderGroupID
    Else
        ShowMe = 0&
    End If

ErrExit:
    Unload Me
    Exit Function

ErrSection:
    Unload Me
    RaiseError "frmOrderGroup.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow the user to close the dialog without saving
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
    RaiseError "frmOrderGroup.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDelete_Click
'' Description: Allow the user to delete an existing order group item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    DeleteOrderGroupItem

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmOrderGroup.cmdDelete_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: Allow the user to edit an existing order group item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    EditOrderGroupItem

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmOrderGroup.cmdEdit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNew_Click
'' Description: Allow the user to create a new order group item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNew_Click()
On Error GoTo ErrSection:

    NewOrderGroupItem

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmOrderGroup.cmdNew_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSave_Click
'' Description: Allow the user to save the order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSave_Click()
On Error GoTo ErrSection:

    If Len(Trim(txtName.Text)) = 0 Then
        MoveFocus txtName
        Err.Raise vbObjectError + 1000, , "Please supply a name for the order group"
    End If
    If Len(Trim(txtName.Text)) > 50 Then
        MoveFocus txtName
        Err.Raise vbObjectError + 1000, , "Please make sure the name is less than 50 characters in length"
    End If

    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroup.cmdSave_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_BeforeMouseDown
'' Description: Show the popup menu if the user right clicks on the grid
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Location of mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid

    With fgOrders
        lMouseRow = .MouseRow
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            .Row = lMouseRow
            .RowSel = lMouseRow
        End If
    End With
    
    EnableControls
    
    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmOrderGroup.fgOrders_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_DblClick
'' Description: Allow the user to edit an item by double clicking on it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    
    With fgOrders
        lMouseRow = .MouseRow
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            .Row = lMouseRow
            .RowSel = lMouseRow
        End If
    End With
    
    EditOrderGroupItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroup.fgOrders_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_KeyUp
'' Description: Perform appropriate action if certain keys are pressed
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Select Case KeyCode
        Case vbKeyDelete
            DeleteOrderGroupItem
            
        Case vbKeyInsert
            NewOrderGroupItem
            
        Case vbKeyReturn
            EditOrderGroupItem
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroup.fgOrders_KeyUp"
    
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

    Dim strPlacement As String          ' Placement for the form
    Dim strFont As String               ' Font from the ini file
    
    g.Styler.StyleForm Me
    
    Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)
    Caption = "Order Group"

    strPlacement = GetIniFileProperty("frmOrderGroup", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement, "LHTW"
    End If
    
    strFont = GetIniFileProperty("frmOrderGroup", "", "Fonts", g.strIniFile)
    If Len(strFont) > 0 Then FontFromString fgOrders.Font, strFont
    
    mnuPopUp.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroup.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Allow only the ShowMe to unload the form
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
    RaiseError "frmOrderGroup.Form_QueryUnload"
    
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

    Dim lMinWidth As Long               ' Minimum scale width
    Dim lMinHeight As Long              ' Minimum scale height
    
    lMinWidth = txtName.Left + txtName.Width + 60
    lMinHeight = fraButtons.Top + fraButtons.Height + 60
    
    If LimitFormSize(Me, lMinWidth, lMinHeight) Then Exit Sub
    
    With txtDescription
        .Move .Left, .Top, ScaleWidth - .Left - 60
    End With
    
    With fraButtons
        .Move ScaleWidth - .Width - 60
    End With
    
    With fgOrders
        .Move 60, .Top, ScaleWidth - fraButtons.Width - 180, ScaleHeight - .Top - 60
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save settings and clean up after ourseleves on an unload
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "frmOrderGroup", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "frmOrderGroup", FontToString(fgOrders.Font), "Fonts", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroup.Form_Unload"
    
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

    With fgOrders
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserResizing = flexResizeNone
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .HighLight = flexHighlightAlways
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_ID)) = "ID"
        .TextMatrix(0, GDCol(eGDCol_OrderText)) = "Order Text"
        .TextMatrix(0, GDCol(eGDCol_DisplayText)) = "Order"
        
        .ColHidden(GDCol(eGDCol_ID)) = True
        .ColHidden(GDCol(eGDCol_OrderText)) = True
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroup.InitGrid"
    
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

    Dim rs As Recordset                 ' Recordset into the database
    Dim Order As New cPtOrder           ' Temporary order object
    
    With fgOrders
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderGroupItems] WHERE [OrderGroupID]=" & Str(m.lOrderGroupID) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = Str(rs!OrderGroupItemID)
            .TextMatrix(.Rows - 1, GDCol(eGDCol_OrderText)) = rs!OrderText
            Set Order = OrderFromOrderText(rs!OrderText)
            .TextMatrix(.Rows - 1, GDCol(eGDCol_DisplayText)) = DisplayTextFromOrder(Order)
            
            rs.MoveNext
        Loop
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroup.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable the controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim bEnable As Boolean              ' Enable or disable controls?
    
    With fgOrders
        If .Row >= .FixedRows And .Row < .Rows Then bEnable = True Else bEnable = False
    End With
    
    Enable cmdEdit, bEnable
    Enable mnuEdit, bEnable
    Enable cmdDelete, bEnable
    Enable mnuDelete, bEnable
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroup.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Long                  ' Index into a for loop
    Dim lID As Long                     ' Order Group Item ID
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderGroups] WHERE [OrderGroupID]=" & Str(m.lOrderGroupID) & ";", dbOpenDynaset)
    If (rs.BOF And rs.EOF) Then
        rs.AddNew
        m.lOrderGroupID = rs!OrderGroupID
    Else
        rs.Edit
    End If
    
    rs!Name = Trim(txtName.Text)
    rs!Description = Trim(txtDescription.Text)
    rs.Update
    
    With fgOrders
        For lIndex = .FixedRows To .Rows - 1
            lID = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_ID))))
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderGroupItems] " & _
                        "WHERE [OrderGroupItemID]=" & Str(lID) & ";", dbOpenDynaset)
            If (rs.BOF And rs.EOF) Then
                rs.AddNew
                .TextMatrix(lIndex, GDCol(eGDCol_ID)) = Str(rs!OrderGroupItemID)
                rs!OrderGroupID = m.lOrderGroupID
            Else
                rs.Edit
            End If
            
            rs!OrderText = .TextMatrix(lIndex, GDCol(eGDCol_OrderText))
            rs.Update
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroup.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuDelete_Click
'' Description: Allow the user to delete an existing order group item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuDelete_Click()
On Error GoTo ErrSection:

    DeleteOrderGroupItem

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmOrderGroup.mnuDelete_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEdit_Click
'' Description: Allow the user to edit an existing order group item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEdit_Click()
On Error GoTo ErrSection:

    EditOrderGroupItem

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmOrderGroup.mnuEdit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuNew_Click
'' Description: Allow the user to create a new order group item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuNew_Click()
On Error GoTo ErrSection:

    NewOrderGroupItem

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmOrderGroup.mnuNew_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtName_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtName_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtName

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmOrderGroup.txtName_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtDescription_GotFocus
'' Description: When the control gets the focus, select all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDescription_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtDescription

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmOrderGroup.txtDescription_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewOrderGroupItem
'' Description: Allow the user to create a new order group item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewOrderGroupItem()
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order object to create
    
    ' Set the default symbol and account before calling the edit order form...
    If Not ActiveChart Is Nothing Then
        Order.SymbolOrSymbolID = RollSymbolForDate(GetSymbol(ActiveChart.SymbolID), Date)
        Order.AccountID = ActiveChart.TradeAccountID
    End If
    
    If frmTTEditOrder.ShowMe(Order, , eGDTTEditOrderMode_FromOrderGroup) = eGDEditOrderReturn_Submit Then
        With fgOrders
            .Redraw = flexRDNone
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = "0"
            .TextMatrix(.Rows - 1, GDCol(eGDCol_OrderText)) = OrderTextFromOrder(Order)
            .TextMatrix(.Rows - 1, GDCol(eGDCol_DisplayText)) = DisplayTextFromOrder(Order)
            
            .Row = .Rows - 1
            .RowSel = .Rows - 1
            
            .AutoSize 0, .Cols - 1, False, 75
            .Redraw = flexRDBuffered
        End With
        
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroup.NewOrderGroupItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditOrderGroupItem
'' Description: Allow the user to edit an existing order group item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditOrderGroupItem()
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order object to create
    Dim lRow As Long                    ' Row in the grid to edit
    
    With fgOrders
        lRow = .Row
        
        If lRow >= .FixedRows And lRow < .Rows Then
            Set Order = OrderFromOrderText(.TextMatrix(lRow, GDCol(eGDCol_OrderText)))
            If frmTTEditOrder.ShowMe(Order, , eGDTTEditOrderMode_FromOrderGroup) = eGDEditOrderReturn_Submit Then
                .TextMatrix(lRow, GDCol(eGDCol_OrderText)) = OrderTextFromOrder(Order)
                .TextMatrix(lRow, GDCol(eGDCol_DisplayText)) = DisplayTextFromOrder(Order)
                .AutoSize 0, .Cols - 1, False, 75
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroup.EditOrderGroupItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteOrderGroupItem
'' Description: Allow the user to delete an existing order group item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteOrderGroupItem()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row in the grid to edit
    
    With fgOrders
        lRow = .Row
        
        If lRow >= .FixedRows And lRow < .Rows Then
            If InfBox("Are you sure that you want to delete this order?", "?", "+Yes|-No", "Confirmation") = "Y" Then
                .RemoveItem lRow
                If lRow - 1 >= .FixedRows And lRow - 1 < .Rows Then
                    .Row = lRow - 1
                    .RowSel = lRow - 1
                End If
                
                EnableControls
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroup.DeleteOrderGroupItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderTextFromOrder
'' Description: Create a string of order text from the order
'' Inputs:      Order
'' Returns:     Order Text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderTextFromOrder(ByVal Order As cPtOrder) As String
On Error GoTo ErrSection:
    
    Dim astrOrder As New cGdArray       ' Array of order text information
    
    With Order
        astrOrder(OrderField(eGDOrderField_Buy)) = Str(.Buy)
        astrOrder(OrderField(eGDOrderField_Quantity)) = Str(.Quantity)
        astrOrder(OrderField(eGDOrderField_Symbol)) = Str(.SymbolOrSymbolID)
        astrOrder(OrderField(eGDOrderField_OrderType)) = Str(.OrderType)
        astrOrder(OrderField(eGDOrderField_StopPrice)) = Str(.StopPrice)
        astrOrder(OrderField(eGDOrderField_LimitPrice)) = Str(.LimitPrice)
        astrOrder(OrderField(eGDOrderField_Account)) = Str(.AccountID)
        astrOrder(OrderField(eGDOrderField_Expiration)) = Str(.Expiration)
    End With
    
    OrderTextFromOrder = astrOrder.JoinFields("|")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOrderGroup.OrderTextFromOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderFromOrderText
'' Description: Create an order from a string of order text
'' Inputs:      Order Text
'' Returns:     Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderFromOrderText(ByVal strOrder As String) As cPtOrder
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order to fill in
    Dim astrOrder As New cGdArray       ' Array of order text information
    
    astrOrder.SplitFields strOrder, "|"
    With Order
        .OrderID = -1&
        .Buy = Val(astrOrder(OrderField(eGDOrderField_Buy)))
        .Quantity = CLng(Val(astrOrder(OrderField(eGDOrderField_Quantity))))
        If Val(astrOrder(OrderField(eGDOrderField_Symbol))) = 0 Then
            .SymbolOrSymbolID = astrOrder(OrderField(eGDOrderField_Symbol))
        Else
            .SymbolOrSymbolID = CLng(Val(astrOrder(OrderField(eGDOrderField_Symbol))))
        End If
        .OrderType = CLng(Val(astrOrder(OrderField(eGDOrderField_OrderType))))
        .StopPrice = Val(astrOrder(OrderField(eGDOrderField_StopPrice)))
        .LimitPrice = Val(astrOrder(OrderField(eGDOrderField_LimitPrice)))
        .AccountID = CLng(Val(astrOrder(OrderField(eGDOrderField_Account))))
        .Expiration = CLng(Val(astrOrder(OrderField(eGDOrderField_Expiration))))
    End With
    
    Set OrderFromOrderText = Order

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOrderGroup.OrderFromOrderText"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayTextFromOrder
'' Description: Create a display string of order text from the order
'' Inputs:      Order
'' Returns:     Display Text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DisplayTextFromOrder(ByVal Order As cPtOrder) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' String to return from the function
    
    strReturn = Order.OrderText
    strReturn = strReturn & " in Account " & g.Broker.AccountNumberForID(Order.AccountID)
    If Order.Expiration = 0 Then
        strReturn = strReturn & " GTC"
    ElseIf Order.Expiration < 0 Then
        strReturn = strReturn & " Day Order"
    Else
        strReturn = strReturn & " Expires " & DateFormat(Order.Expiration)
    End If
    
    DisplayTextFromOrder = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOrderGroup.DisplayTextFromOrder"
    
End Function

