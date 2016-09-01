VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmCattleManage 
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   2295
      Left            =   3300
      TabIndex        =   1
      Top             =   180
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
      Caption         =   "frmCattleManage.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCattleManage.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCattleManage.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   0
         TabIndex        =   6
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
         Caption         =   "frmCattleManage.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCattleManage.frx":0096
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCattleManage.frx":00B6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRemove 
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   1800
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
         Caption         =   "frmCattleManage.frx":00D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCattleManage.frx":0100
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCattleManage.frx":0120
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   1260
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
         Caption         =   "frmCattleManage.frx":013C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCattleManage.frx":0166
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCattleManage.frx":0186
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   720
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
         Caption         =   "frmCattleManage.frx":01A2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCattleManage.frx":01CA
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCattleManage.frx":01EA
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdClose 
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
         Caption         =   "frmCattleManage.frx":0206
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCattleManage.frx":0232
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCattleManage.frx":0252
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgObjects 
      Height          =   2895
      Left            =   180
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
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
   End
End
Attribute VB_Name = "frmCattleManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmCattleManage.frm
'' Description: Form for allowing user to manage certain Turnkey stuff
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/25/2014   DAJ         Rations/Ingredients
'' 03/07/2014   DAJ         Moved into NavCattle.DLL
'' 03/20/2014   DAJ         Added a "Click Here" line
'' 04/15/2014   DAJ         Don't allow user to delete ingredient used in details
'' 05/22/2014   DAJ         Renamed frmTurnkeyManage to frmCattleManage
'' 05/22/2014   DAJ         Renamed frmTurnkeyEditor to frmCattleEditor; Renamed g.Turnkey to g.Cattle
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDManageModes
    eGDManageMode_FeedYards
    eGDManageMode_FeedYardCustomers
    eGDManageMode_Ingredients
    eGDManageMode_Rations
End Enum

Private Enum eGDIngredientCols
    eGDIngredientCols_IngredientID = 0
    eGDIngredientCols_Ingredient = 1
    eGDIngredientCols_CostPerPound = 2
    eGDIngredientCols_DryFeedPct = 3
    
    eGDIngredientCols_NumCols
End Enum

Private Type mPrivate
    nMode As eGDManageModes             ' Mode of the form
    bOK As Boolean                      ' Did the user click OK?
    bClosing As Boolean                 ' Is the form closing?
    bSelectMode As Boolean              ' Select mode?
    
    strDryFeedPct As String             ' Default dry feed percent
    strFeedyardID As String             ' Feed Yard ID
    Ration As cBrokerMessage            ' Ration object
    strName As String                   ' Name of the object
    iButton As Integer                  ' Mouse button pressed

    lExtendCol As Long                  ' Extend column
    lPrevColWidth As Long               ' Previous column width
End Type
Private m As mPrivate

Private Property Get IngredientCol(ByVal nCol As eGDIngredientCols) As Long
    IngredientCol = nCol
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeFeedYards
'' Description: Setup and show form for managing feed yards
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMeFeedYards()
On Error GoTo ErrSection:

    m.nMode = eGDManageMode_FeedYards
    Caption = g.Cattle.ProductName & " Feedyards"
    m.bClosing = False
    m.strName = ""
    m.bSelectMode = False
    
    fraButtons.Visible = True
    cmdClose.Caption = "&Close"
    cmdCancel.Visible = False
    cmdEdit.Visible = True
    cmdAdd.Top = 720

    InitGridFeedYards
    LoadGridFeedYards

    ShowForm Me, eForm_Modal, g.frmMain

ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmCattleManage.ShowMeFeedYards"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeFeedYardCustomers
'' Description: Setup and show form for managing feed yard customers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMeFeedYardCustomers()
On Error GoTo ErrSection:

    m.nMode = eGDManageMode_FeedYardCustomers
    Caption = g.Cattle.ProductName & " Feedyard Customers"
    m.bClosing = False
    m.strName = ""
    m.bSelectMode = False
    
    fraButtons.Visible = True
    cmdClose.Caption = "&Close"
    cmdCancel.Visible = False
    cmdEdit.Visible = True
    cmdAdd.Top = 720

    InitGridFeedYardCustomers
    LoadGridFeedYardCustomers

    ShowForm Me, eForm_Modal, g.frmMain

ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmCattleManage.ShowMeFeedYardCustomers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeIngredients
'' Description: Setup and show form for managing ingredients
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMeIngredients()
On Error GoTo ErrSection:

    m.nMode = eGDManageMode_Ingredients
    Caption = g.Cattle.ProductName & " Ingredients"
    m.bClosing = False
    m.strName = ""
    m.bSelectMode = False
    
    fraButtons.Visible = True
    cmdClose.Caption = "&OK"
    cmdCancel.Visible = True
    cmdEdit.Visible = False
    cmdAdd.Top = cmdEdit.Top
    
    InitGridIngredients
    LoadGridIngredients

    ShowForm Me, eForm_Modal, g.frmMain
    
    If m.bOK Then
        SaveIngredients
    End If

ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmCattleManage.ShowMeIngredients"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeRations
'' Description: Setup and show form for managing rations
'' Inputs:      Select Mode?
'' Returns:     Selected Ration ( Nothing if not selected )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeRations(Optional ByVal bSelectMode As Boolean = False) As cBrokerMessage
On Error GoTo ErrSection:

    Dim Ration As cBrokerMessage        ' Selected ration to remove

    m.nMode = eGDManageMode_Rations
    Caption = g.Cattle.ProductName & " Rations"
    m.bClosing = False
    m.strName = ""
    m.bSelectMode = bSelectMode
    
    fraButtons.Visible = True
    If bSelectMode = True Then
        cmdClose.Caption = "&OK"
        cmdCancel.Visible = True
        fraButtons.Height = fraButtons.Height + cmdRemove.Height
        cmdAdd.Top = cmdEdit.Top
        cmdEdit.Top = cmdRemove.Top
        cmdRemove.Top = cmdEdit.Top + cmdEdit.Height
    Else
        cmdClose.Caption = "&Close"
        cmdCancel.Visible = False
        cmdAdd.Top = 720
    End If
    cmdEdit.Visible = True
    
    InitGridRations
    LoadGridRations

    ShowForm Me, eForm_Modal, g.frmMain
    
    Set Ration = Nothing
    With fgObjects
        If (bSelectMode = True) And (m.bOK = True) Then
            If ValidGridRow(fgObjects) Then
                Set Ration = .RowData(.Row)
            End If
        End If
    End With
    
    Set ShowMeRations = Ration

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmCattleManage.ShowMeRations"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_FeedYard
'' Description: Feed yard record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_FeedYard(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim FeedYard As cBrokerMessage      ' Feed Yard object
    Dim lRow As Long                    ' Row in the grid
    Dim strFirstField As String         ' First field in the message

    If m.bClosing = False Then
        If (m.nMode = eGDManageMode_FeedYards) And (Len(strMessage) > 0) Then
            strFirstField = Parse(strMessage, vbTab, 1)
            
            If UCase(strFirstField) = "BEGIN" Then
                fgObjects.Rows = fgObjects.FixedRows
            ElseIf UCase(strFirstField) = "END" Then
            Else
                Set FeedYard = New cBrokerMessage
                FeedYard.FromString strMessage
                
                lRow = FindRow(FeedYard("ID"))
                FeedyardToGrid FeedYard, lRow
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.Turnkey_FeedYard", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_Customer
'' Description: Customer record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_Customer(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim Customer As cBrokerMessage      ' Customer object
    Dim lRow As Long                    ' Row in the grid
    Dim strFirstField As String         ' First field in the message

    If m.bClosing = False Then
        If (m.nMode = eGDManageMode_FeedYardCustomers) And (Len(strMessage) > 0) Then
            strFirstField = Parse(strMessage, vbTab, 1)
            
            If UCase(strFirstField) = "BEGIN" Then
                fgObjects.Rows = fgObjects.FixedRows
            ElseIf UCase(strFirstField) = "END" Then
            Else
                Set Customer = New cBrokerMessage
                Customer.FromString strMessage
                
                lRow = FindRow(Customer("ID"))
                FeedyardCustomerToGrid Customer, lRow
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.Turnkey_Customer", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_Ingredient
'' Description: Ingredient record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_Ingredient(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim Ingredient As cBrokerMessage    ' Ingredient object
    Dim lRow As Long                    ' Row in the grid
    Dim strFirstField As String         ' First field in the message

    If m.bClosing = False Then
        If (m.nMode = eGDManageMode_Ingredients) And (Len(strMessage) > 0) Then
            strFirstField = Parse(strMessage, vbTab, 1)
            
            If UCase(strFirstField) = "BEGIN" Then
                fgObjects.Rows = fgObjects.FixedRows
            ElseIf UCase(strFirstField) = "END" Then
            Else
                Set Ingredient = New cBrokerMessage
                Ingredient.FromString strMessage
                
                lRow = FindRow(Ingredient("ID"))
                IngredientToGrid Ingredient, lRow
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.Turnkey_Ingredient", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_Ration
'' Description: Ration record returned from the Turnkey source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_Ration(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim Ration As cBrokerMessage        ' Ration object
    Dim lRow As Long                    ' Row in the grid
    Dim strFirstField As String         ' First field in the message

    If (m.bClosing = False) And (Len(strMessage) > 0) Then
        Select Case m.nMode
            Case eGDManageMode_Rations
                strFirstField = Parse(strMessage, vbTab, 1)
                
                If UCase(strFirstField) = "BEGIN" Then
                    fgObjects.Rows = fgObjects.FixedRows
                ElseIf UCase(strFirstField) = "END" Then
                Else
                    Set Ration = New cBrokerMessage
                    Ration.FromString strMessage
                    
                    lRow = FindRow(Ration("ID"))
                    RationToGrid Ration, lRow
                End If
        
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.Turnkey_Ration", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAdd_Click
'' Description: Add a new object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAdd_Click()
On Error GoTo ErrSection:

    AddObject

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.cmdAdd_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Cancel the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bClosing = True
    m.bOK = False
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.cmdCancel_Click"
    
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

    If Validate Then
        m.bClosing = True
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.cmdClose_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: Edit an existing object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    EditObject

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.cmdEdit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemove_Click
'' Description: Remove an existing object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemove_Click()
On Error GoTo ErrSection:

    RemoveObject

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.cmdRemove_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_AfterEdit
'' Description: Handle the user changing the value of a cell
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim Ingredient As cBrokerMessage    ' Ingredient object

    If m.nMode = eGDManageMode_Ingredients Then
        With fgObjects
            If TypeOf .RowData(Row) Is cBrokerMessage Then
                Set Ingredient = .RowData(Row)
                
                Select Case Col
                    Case IngredientCol(eGDIngredientCols_Ingredient):
                        Ingredient.Add "Ingredient", .TextMatrix(Row, Col)
                    Case IngredientCol(eGDIngredientCols_CostPerPound):
                        Ingredient.Add "CostPerPound", .TextMatrix(Row, Col)
                    Case IngredientCol(eGDIngredientCols_DryFeedPct):
                        Ingredient.Add "DryFeedPct", .TextMatrix(Row, Col)
                End Select
                
                .RowData(Row) = Ingredient
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.fgObjects_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_AfterRowColChange
'' Description: Handle the user moving cells
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If m.nMode = eGDManageMode_Ingredients Then
        EditCell fgObjects
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.fgObjects_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_BeforeEdit
'' Description: Decide if we want the cell to be edited
'' Inputs:      Row, Column, Cancel Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Select Case m.nMode
        Case eGDManageMode_Ingredients
            Cancel = RowIsClickHereLine(Row)
        
        Case Else
            Cancel = True
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.fgObjects_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_BeforeMouseDown
'' Description: Bring up the context menu on a right-click in the grid
'' Inputs:      Mouse Button pressed, Shift/Ctrl/Alt status, Mouse location,
''              Cancel the Mouse Down?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Mouse row in the grid

    m.iButton = Button
    If Button = vbRightButton Then
        With fgObjects
            lMouseRow = .MouseRow
            
            .Row = lMouseRow
            
            If Validate(False) Then
                Select Case m.nMode
                    Case eGDManageMode_FeedYards
                        mnuAdd.Caption = "Add Feedyard"
                        mnuEdit.Visible = True
                        mnuEdit.Caption = "Edit Feedyard"
                        mnuRemove.Caption = "Remove Feedyard"
                
                    Case eGDManageMode_FeedYardCustomers
                        mnuAdd.Caption = "Add Feedyard Customer"
                        mnuEdit.Visible = True
                        mnuEdit.Caption = "Edit Feedyard Customer"
                        mnuRemove.Caption = "Remove Feedyard Customer"
                
                    Case eGDManageMode_Ingredients
                        mnuAdd.Caption = "Add Ingredient"
                        mnuEdit.Visible = False
                        mnuEdit.Caption = "Edit Ingredient"
                        mnuRemove.Caption = "Remove Ingredient"
                
                    Case eGDManageMode_Rations
                        mnuAdd.Caption = "Add Ration"
                        mnuEdit.Visible = True
                        mnuEdit.Caption = "Edit Ration"
                        mnuRemove.Caption = "Remove Ration"
                
                End Select
                
                Enable mnuEdit, ValidGridRow(fgObjects, lMouseRow)
                Enable mnuRemove, mnuEdit.Enabled
                
                PopupMenu mnuPopup
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.fgObjects_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_BeforeRowColChange
'' Description: Handle the user wanting to change the current cell
'' Inputs:      Old Row and Column, New Row and Column, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If (NewRow <> OldRow) And (OldRow <> -1&) Then
        Select Case m.nMode
            Case eGDManageMode_Ingredients
                If Len(fgObjects.TextMatrix(OldRow, IngredientCol(eGDIngredientCols_Ingredient))) = 0 Then
                    InfBox "You must select an ingredient", "!", , "Error"
                    Cancel = True
                    
                    EditCell fgObjects, , IngredientCol(eGDIngredientCols_Ingredient)
                End If
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.fgObjects_BeforeRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_Click
'' Description: Handle a user click in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_Click()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Mouse row in the grid
    
    If m.iButton = vbLeftButton Then
        If m.nMode = eGDManageMode_Ingredients Then
            With fgObjects
                lMouseRow = .MouseRow
                
                If RowIsClickHereLine(lMouseRow) Then
                    AddObject
                End If
            End With
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.fgObjects_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_DblClick
'' Description: If the user double clicks on an object, edit it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Mouse row in the grid
    
    If m.iButton = vbLeftButton Then
        With fgObjects
            lMouseRow = .MouseRow
            
            If ValidGridRow(fgObjects, lMouseRow) Then
                .Row = lMouseRow
                
                If m.bSelectMode = True Then
                    m.bOK = True
                    Hide
                Else
                    EditObject
                End If
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.fgObjects_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_KeyDown
'' Description: Handle the user pressing the Insert or Delete keys in the grid
'' Inputs:      Key Pressed, Shift/Ctrl/Alt Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyInsert Then
        AddObject
    ElseIf KeyCode = vbKeyDelete Then
        RemoveObject
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.fgObjects_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_KeyPress
'' Description: Handle the user pressing the Enter key in the grid
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        If m.bSelectMode = True Then
            m.bOK = True
            Hide
        Else
            EditObject
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.fgObjects_KeyPress"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_ValidateEdit
'' Description: Validate the information entered by the user
'' Inputs:      Row, Column, Cancel Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim dValue As Double                ' Value of the edit text

    If (m.bClosing = False) Or (m.bOK = True) Then
        With fgObjects
            If m.nMode = eGDManageMode_Ingredients Then
                Select Case Col
                    Case IngredientCol(eGDIngredientCols_Ingredient):
                        'If Len(.EditText) = 0 Then
                        '    InfBox "You must enter in a name for the ingredient", "!", , "Error"
                        '    Cancel = True
                        'End If
                        
                    Case IngredientCol(eGDIngredientCols_CostPerPound):
                        If IsAlpha(.EditText) Then
                            InfBox "Cost value must be a number", "!", , "Error"
                            Cancel = True
                        End If
                        
                    Case IngredientCol(eGDIngredientCols_DryFeedPct):
                        If IsAlpha(.EditText) Then
                            InfBox "Dry Feed Percent must be a number", "!", , "Error"
                            Cancel = True
                        Else
                            dValue = ValOfText(.EditText)
                            If (dValue < 0) Or (dValue > 100) Then
                                InfBox "Dry Feed Percent must be between 0 and 100", "!", , "Error"
                                Cancel = True
                            End If
                        End If
                
                End Select
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.fgObjects_ValidateEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = g.AppBridge.Picture16(g.Cattle.IconName)
    
    g.Styler.StyleForm Me
    
    PlaceForm Me
    
    mnuPopup.Visible = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Determine whether or not to let the form close
'' Inputs:      Cancel Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        m.bClosing = True
        m.bOK = False
        Hide
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCattleManage.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Size and move controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lSpace As Long                  ' Space between the controls
    Dim lMinScaleWidth As Long          ' Minimum scale width for the form
    Dim lMinScaleHeight As Long         ' Minimum scale height for the form
    
    lSpace = 120
    lMinScaleWidth = (fraButtons.Width * 3) + (lSpace * 3)
    lMinScaleHeight = fraButtons.Height + (lSpace * 2)
    
    If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) = False Then
        With fraButtons
            .Move ScaleWidth - lSpace - .Width, lSpace
        End With
        
        With fgObjects
            .Move lSpace, lSpace, fraButtons.Left - (lSpace * 2), ScaleHeight - (lSpace * 2)
        End With
        
        ExtendCustomColumn
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
    RaiseError "frmCattleManage.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuAdd_Click
'' Description: Handle the user wanting to add an object from the context menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuAdd_Click()
On Error GoTo ErrSection:

    AddObject

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.mnuAdd_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEdit_Click
'' Description: Handle the user wanting to edit an object from the context menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEdit_Click()
On Error GoTo ErrSection:

    EditObject

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.mnuEdit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuRemove_Click
'' Description: Handle the user wanting to remove an object from the context menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRemove_Click()
On Error GoTo ErrSection:

    RemoveObject

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.mnuRemove_Click"
    
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
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = True
        .BackColorAlternate = .BackColor
        .BackColorBkg = vbApplicationWorkspace
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridInset
        .MergeCells = flexMergeNever
        .OutlineBar = flexOutlineBarNone
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .WordWrap = True
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGridFeedYards
'' Description: Initialize the grid for feed yards
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGridFeedYards()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        InitGrid
        
        .Rows = 1
        .FixedRows = 1
        .Cols = 3
        .FixedCols = 0
        
        .TextMatrix(0, 0) = "Name"
        .TextMatrix(0, 1) = "Code"
        .TextMatrix(0, 2) = "Dry Feed Pct"
        
        m.lExtendCol = 0
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.InitGridFeedYards"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGridFeedYardCustomers
'' Description: Initialize the grid for feed yard customers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGridFeedYardCustomers()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        InitGrid
        
        .Rows = 1
        .FixedRows = 1
        .Cols = 2
        .FixedCols = 0
        
        .TextMatrix(0, 0) = "Name"
        .TextMatrix(0, 1) = "Number"
        
        m.lExtendCol = 0
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.InitGridFeedYardCustomers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGridIngredients
'' Description: Initialize the grid for ingredients
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGridIngredients()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        InitGrid
        
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .MergeCells = flexMergeFree
        .SelectionMode = flexSelectionFree
        .TabBehavior = flexTabCells
        
        .Rows = 1
        .FixedRows = 1
        .Cols = IngredientCol(eGDIngredientCols_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, IngredientCol(eGDIngredientCols_IngredientID)) = "ID"
        .TextMatrix(0, IngredientCol(eGDIngredientCols_Ingredient)) = "Name"
        .TextMatrix(0, IngredientCol(eGDIngredientCols_CostPerPound)) = "Cost ($/lb)"
        .TextMatrix(0, IngredientCol(eGDIngredientCols_DryFeedPct)) = "Dry Feed Pct"
        
        .ColHidden(IngredientCol(eGDIngredientCols_IngredientID)) = True
        .ColFormat(IngredientCol(eGDIngredientCols_CostPerPound)) = "$#,##0.000"
        
        m.lExtendCol = IngredientCol(eGDIngredientCols_Ingredient)
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.InitGridIngredients"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGridRations
'' Description: Initialize the grid for rations
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGridRations()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        InitGrid
        
        .BackColorBkg = vbWindowBackground
        .GridLines = flexGridNone
        .GridLinesFixed = flexGridNone
        
        .Rows = 0
        .FixedRows = 0
        .Cols = 1
        .FixedCols = 0
        
        m.lExtendCol = 0
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.InitGridRations"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGridRation
'' Description: Initialize the grid for a ration
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGridRation()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        InitGrid
        
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .SelectionMode = flexSelectionFree
        .TabBehavior = flexTabCells
        
        .Rows = 1
        .FixedRows = 1
        .Cols = 4
        .FixedCols = 0
        
        .TextMatrix(0, 0) = "IngredientID"
        .TextMatrix(0, 1) = "Ingredient"
        .TextMatrix(0, 2) = "Pounds Fed"
        .TextMatrix(0, 3) = "% Markup"
        
        .ColHidden(0) = True
        
        m.lExtendCol = 1
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.InitGridRation"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridFeedYards
'' Description: Load the grid for feed yards
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridFeedYards()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim FeedYards As cGdTree            ' Feed yard collection
    Dim lIndex As Long                  ' Index into a for loop
    
    Set FeedYards = g.Cattle.FeedYards
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 1 To FeedYards.Count
            FeedyardToGrid FeedYards(lIndex)
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.LoadGridFeedYards"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridFeedYardCustomers
'' Description: Load the grid for feed yard customers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridFeedYardCustomers()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim Customers As cGdTree            ' Feed yard customer collection
    Dim lIndex As Long                  ' Index into a for loop
    
    Set Customers = g.Cattle.Customers
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 1 To Customers.Count
            FeedyardCustomerToGrid Customers(lIndex)
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.LoadGridFeedYardCustomers"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridIngredients
'' Description: Load the grid for ingredients
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridIngredients()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim Ingredients As cGdTree          ' Ingredients collection
    Dim lIndex As Long                  ' Index into a for loop
    Dim strLotColumnID As String        ' Lot Column ID
    Dim strDetailOptions As String      ' Detail options
    Dim astrDetailOptions As cGdArray   ' Detail options split out into an array
    Dim Ingredient As cBrokerMessage    ' Ingredient object
    Dim FeedYard As cBrokerMessage      ' Selected feedyard
    
    Set Ingredients = g.Cattle.Ingredients
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        Set FeedYard = g.Cattle.SelectedFeedYard
        If Not FeedYard Is Nothing Then
            m.strDryFeedPct = FeedYard("DryFeedPct")
            m.strFeedyardID = FeedYard("ID")
        Else
            m.strDryFeedPct = ""
            m.strFeedyardID = ""
        End If
        
        .Rows = .FixedRows
        If Ingredients.Count > 0 Then
            For lIndex = 1 To Ingredients.Count
                IngredientToGrid Ingredients(lIndex)
            Next lIndex
        Else
            If g.Cattle.LotColumnMap.Exists("Ingredient") Then
                strLotColumnID = g.Cattle.LotColumnMap("Ingredient")
                strDetailOptions = g.Cattle.DetailOptions(strLotColumnID)
                If Len(strDetailOptions) > 0 Then
                    Set astrDetailOptions = New cGdArray
                    astrDetailOptions.SplitFields strDetailOptions, "|"
                    
                    For lIndex = 0 To astrDetailOptions.Size - 1
                        Set Ingredient = New cBrokerMessage
                        
                        Ingredient.Add "ID", ""
                        Ingredient.Add "FeedYardID", m.strFeedyardID
                        Ingredient.Add "Ingredient", astrDetailOptions(lIndex)
                        Ingredient.Add "CostPerPound", ""
                        Ingredient.Add "DryFeedPct", m.strDryFeedPct
                        
                        IngredientToGrid Ingredient
                    Next lIndex
                End If
            End If
        End If
        
        AddClickHereLine
        
        .AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.LoadGridIngredients"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridRations
'' Description: Load the grid for rations
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridRations()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim Rations As cGdTree              ' Rations collection
    Dim lIndex As Long                  ' Index into a for loop
    
    Set Rations = g.Cattle.Rations
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        For lIndex = 1 To Rations.Count
            RationToGrid Rations(lIndex)
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.LoadGridRations"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridRation
'' Description: Load the grid for a ration
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridRation()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim astrIngredient As cGdArray      ' Ingredient information split out into an array
    Dim astrPoundsFed As cGdArray       ' Pounds Fed information split out into an array
    Dim astrPctMarkup As cGdArray       ' Percent Markup informaion split out into an array
    Dim lIndex As Long                  ' Index into a for loop
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        Set astrIngredient = New cGdArray
        Set astrPoundsFed = New cGdArray
        Set astrPctMarkup = New cGdArray
        
        astrIngredient.SplitFields m.Ration("IngredientID"), "|"
        astrPoundsFed.SplitFields m.Ration("PoundsFed"), "|"
        astrPctMarkup.SplitFields m.Ration("PercentMarkup"), "|"
        
        .Rows = astrIngredient.Size + .FixedRows
        
        For lIndex = 0 To astrIngredient.Size - 1
            .TextMatrix(lIndex + .FixedRows, 0) = astrIngredient(lIndex)
            .TextMatrix(lIndex + .FixedRows, 1) = g.Cattle.IngredientNameForID(astrIngredient(lIndex))
            .TextMatrix(lIndex + .FixedRows, 2) = astrPoundsFed(lIndex)
            .TextMatrix(lIndex + .FixedRows, 3) = astrPctMarkup(lIndex)
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.LoadGridRation"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtendCustomColumn
'' Description: Adjust all column widths to accomodate custom extended column
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExtendCustomColumn(Optional ByVal nResizeCol As Long = -1)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lTotal As Long                  ' Total width
    Dim lDiff As Long                   ' Difference in column width

    With fgObjects
        ' if column being resized is after the extended column,
        ' then change the width of the next visible column instead
        If nResizeCol >= m.lExtendCol Then
            .Redraw = flexRDNone
            lDiff = .ColWidth(nResizeCol) - m.lPrevColWidth
            For lIndex = nResizeCol + 1 To .Cols - 1
                If Not .ColHidden(lIndex) Then
                    .ColWidth(lIndex) = .ColWidth(lIndex) - lDiff
                    Exit For
                End If
            Next
            m.lPrevColWidth = 0
        End If
        
        ' size the custom extended column in order to fill the client width
        .ColHidden(m.lExtendCol) = True
        .Redraw = flexRDBuffered '(must do this so .ClientWidth will be correct)
        .Redraw = flexRDNone
        lTotal = 0
        For lIndex = 0 To .Cols - 1
            If Not .ColHidden(lIndex) Then
                lTotal = lTotal + .ColWidth(lIndex)
            End If
        Next
        lTotal = .ClientWidth - lTotal
        If lTotal > 0 Then .ColWidth(m.lExtendCol) = lTotal
        .ColHidden(m.lExtendCol) = False
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.ExtendCustomColumn"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FeedYardToGrid
'' Description: Put the given feed yard in the grid
'' Inputs:      Feed Yard, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FeedyardToGrid(ByVal FeedYard As cBrokerMessage, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = FeedYard
        .TextMatrix(lRow, 0) = FeedYard("Name")
        .TextMatrix(lRow, 1) = FeedYard("Code")
        .TextMatrix(lRow, 2) = FeedYard("DryFeedPct")
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.FeedYardToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FeedYardCustomerToGrid
'' Description: Put the given feed yard customer in the grid
'' Inputs:      Feed Yard Customer, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FeedyardCustomerToGrid(ByVal FeedYardCustomer As cBrokerMessage, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = FeedYardCustomer
        .TextMatrix(lRow, 0) = FeedYardCustomer("Name")
        .TextMatrix(lRow, 1) = FeedYardCustomer("Number")
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.FeedYardCustomerToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IngredientToGrid
'' Description: Put the given ingredient in the grid
'' Inputs:      Ingredient, Row
'' Returns:     Row
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IngredientToGrid(ByVal Ingredient As cBrokerMessage, Optional ByVal lRow As Long = -1&) As Long
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lClickHereLine As Long          ' Click here row
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            
            lClickHereLine = ClickHereLine
            If lClickHereLine = -1& Then
                lRow = .Rows - 1
            Else
                .RowPosition(.Rows - 1) = lClickHereLine
                lRow = lClickHereLine
            End If
        End If
        
        .MergeRow(lRow) = False
        .RowData(lRow) = Ingredient
        .TextMatrix(lRow, IngredientCol(eGDIngredientCols_IngredientID)) = Ingredient("ID")
        .TextMatrix(lRow, IngredientCol(eGDIngredientCols_Ingredient)) = Ingredient("Ingredient")
        .TextMatrix(lRow, IngredientCol(eGDIngredientCols_CostPerPound)) = Ingredient("CostPerPound")
        .TextMatrix(lRow, IngredientCol(eGDIngredientCols_DryFeedPct)) = Ingredient("DryFeedPct")
        
        .Redraw = nRedraw
    End With

    IngredientToGrid = lRow

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleManage.IngredientToGrid"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RationToGrid
'' Description: Put the given ration in the grid
'' Inputs:      Ration, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RationToGrid(ByVal Ration As cBrokerMessage, Optional ByVal lRow As Long = -1&)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
        
        .RowData(lRow) = Ration
        .TextMatrix(lRow, 0) = Ration("RationName")
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.RationToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveIngredients
'' Description: Save the ingredients
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveIngredients()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Ingredient As cBrokerMessage    ' Ingredient object
    Dim strRemoved As String            ' Ingredient ID's removed
    
    With fgObjects
        strRemoved = ""
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cBrokerMessage Then
                Set Ingredient = .RowData(lIndex)
                
                If Ingredient("Deleted") = "1" Then
                    strRemoved = strRemoved & Ingredient("ID") & "|"
                End If
            End If
        Next lIndex
        
        If Len(strRemoved) > 1 Then
            strRemoved = Left(strRemoved, Len(strRemoved) - 1)
            g.Cattle.RemoveIngredientFromRations strRemoved
        End If
        
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cBrokerMessage Then
                Set Ingredient = .RowData(lIndex)
                
                If Not g.Cattle Is Nothing Then
                    g.Cattle.UpdateIngredient Ingredient
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.SaveIngredients"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedObject
'' Description: Grab the object on the selected row in the grid
'' Inputs:      None
'' Returns:     Selected Object ( Nothing if none )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectedObject() As cBrokerMessage
On Error GoTo ErrSection:

    Dim ReturnObject As cBrokerMessage  ' Object to return
    
    Set ReturnObject = Nothing
    With fgObjects
        If ValidGridRow(fgObjects) Then
            Set ReturnObject = .RowData(.Row)
        End If
    End With
    
    Set SelectedObject = ReturnObject

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleManage.SelectedObject"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddObject
'' Description: Add a new object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddObject()
On Error GoTo ErrSection:

    Dim Ingredient As cBrokerMessage    ' Ingredient object
    Dim Ration As cBrokerMessage        ' Ration object
    Dim lRow As Long                    ' Row in the grid

    Select Case m.nMode
        Case eGDManageMode_FeedYards
            g.Cattle.UpdateFeedYard
            
        Case eGDManageMode_FeedYardCustomers
            g.Cattle.UpdateFeedYardCustomer
            
        Case eGDManageMode_Ingredients
            With fgObjects
                Set Ingredient = New cBrokerMessage
                Ingredient.Add "ID", ""
                Ingredient.Add "FeedYardID", m.strFeedyardID
                Ingredient.Add "Ingredient", ""
                Ingredient.Add "CostPerPound", ""
                Ingredient.Add "DryFeedPct", m.strDryFeedPct
                
                lRow = IngredientToGrid(Ingredient)
                EditCell fgObjects, lRow, IngredientCol(eGDIngredientCols_Ingredient)
            End With
            
        Case eGDManageMode_Rations
            Set Ration = New cBrokerMessage
            
            Ration.Add "ID", ""
            Ration.Add "RationName", ""
            Ration.Add "IngredientID", ""
            Ration.Add "PoundsFed", ""
            Ration.Add "PercentMarkup", ""
            
            frmCattleEditor.ShowMeRation Ration
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.AddObject"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditObject
'' Description: Edit an existing object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditObject()
On Error GoTo ErrSection:

    Dim ToEdit As cBrokerMessage        ' Object to edit

    Set ToEdit = SelectedObject
    If Not SelectedObject Is Nothing Then
        Select Case m.nMode
            Case eGDManageMode_FeedYards
                g.Cattle.UpdateFeedYard ToEdit
                
            Case eGDManageMode_FeedYardCustomers
                g.Cattle.UpdateFeedYardCustomer ToEdit
                
            Case eGDManageMode_Rations
                frmCattleEditor.ShowMeRation ToEdit
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.EditObject"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveObject
'' Description: Remove an existing object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveObject()
On Error GoTo ErrSection:

    Dim ToRemove As cBrokerMessage      ' Object to remove
    Dim bInRation As Boolean            ' Is an ingredient in a ration?
    Dim strResponse As String           ' Response from the InfBox

    Set ToRemove = SelectedObject
    If Not SelectedObject Is Nothing Then
        Select Case m.nMode
            Case eGDManageMode_FeedYards
                g.Cattle.RemoveFeedYard ToRemove
                
            Case eGDManageMode_FeedYardCustomers
                g.Cattle.RemoveFeedYardCustomer ToRemove
                
            Case eGDManageMode_Ingredients
                If Len(ToRemove("ID")) = 0 Then
                    fgObjects.RemoveItem fgObjects.Row
                Else
                    If g.Cattle.IngredientUsedInDetails(ToRemove("Ingredient")) = True Then
                        InfBox "You cannot remove '" & ToRemove("Ingredient") & "' because it is used in lots", "!", , "Error"
                    Else
                        bInRation = g.Cattle.IngredientUsedInRations(ToRemove("ID"))
                        
                        If bInRation Then
                            strResponse = InfBox("Removing '" & ToRemove("Ingredient") & "' will result in the ingredient being removed from one or more rations.||Do you want to continue?", "?", "+Yes|-No", "Confirmation")
                        Else
                            strResponse = InfBox("Are you sure you want to remove '" & ToRemove("Ingredient") & "'?", "?", "+Yes|-No", "Confirmation")
                        End If
                        
                        If strResponse = "Y" Then
                            ToRemove.Add "Deleted", "1"
                            
                            ' This will get done at the time of the save...
                            'g.Cattle.UpdateIngredient ToRemove
                            
                            fgObjects.RowData(fgObjects.Row) = ToRemove
                            fgObjects.RowHidden(fgObjects.Row) = True
                        End If
                    End If
                End If
                SetBackColors fgObjects
                
            Case eGDManageMode_Rations
                If Len(ToRemove("ID")) = 0 Then
                    fgObjects.RemoveItem fgObjects.Row
                Else
                    If InfBox("Are you sure you want to remove '" & ToRemove("RationName") & "'?", "?", "+Yes|-No", "Confirmation") = "Y" Then
                        ToRemove.Add "Deleted", "1"
                        g.Cattle.UpdateRation ToRemove
                        
                        fgObjects.RowData(fgObjects.Row) = ToRemove
                        
                        fgObjects.RowHidden(fgObjects.Row) = True
                    End If
                End If
                SetBackColors fgObjects
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.RemoveObject"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FindRow
'' Description: Find the row for the given ID
'' Inputs:      ID
'' Returns:     Row where that ID exists ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindRow(ByVal strID As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim RowObject As cBrokerMessage     ' Row object
    
    lReturn = -1&
    With fgObjects
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cBrokerMessage Then
                Set RowObject = .RowData(lIndex)
                If RowObject("ID") = strID Then
                    lReturn = lIndex
                    Exit For
                End If
            End If
        Next lIndex
    End With
    
    FindRow = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleManage.FindRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Validate
'' Description: Validate the information in the grid
'' Inputs:      Show Message?
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Validate(Optional ByVal bShowMessage As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = True
    With fgObjects
        Select Case m.nMode
            Case eGDManageMode_Ingredients
                For lIndex = .FixedRows To .Rows - 1
                    If Len(.TextMatrix(lIndex, IngredientCol(eGDIngredientCols_Ingredient))) = 0 Then
                        If bShowMessage Then
                            InfBox "You must enter a name for the ingredient", "!", , "Error"
                            EditCell fgObjects, lIndex, IngredientCol(eGDIngredientCols_Ingredient)
                        End If
                        
                        bReturn = False
                        Exit For
                    End If
                Next lIndex
        
        End Select
    End With
    
    Validate = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleManage.Validate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClickHereLine
'' Description: Determine the row in the grid that is the click here line
'' Inputs:      None
'' Returns:     Row of the click here line ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ClickHereLine() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop

    lReturn = -1&
    If m.nMode = eGDManageMode_Ingredients Then
        With fgObjects
            For lIndex = .FixedRows To .Rows - 1
                If RowIsClickHereLine(lIndex) Then
                    lReturn = lIndex
                    Exit For
                End If
            Next lIndex
        End With
    End If

    ClickHereLine = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleManage.ClickHereLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddClickHereLine
'' Description: Add the click here line if it doesn't exist
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddClickHereLine()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim strText As String               ' Text to put in the grid
    Dim lStartCol As Long               ' Starting column
    Dim lEndCol As Long                 ' Ending column

    If m.nMode = eGDManageMode_Ingredients Then
        If ClickHereLine = -1& Then
            With fgObjects
                nRedraw = .Redraw
                .Redraw = flexRDNone
                
                .Rows = .Rows + 1
                .MergeRow(.Rows - 1) = True
                
                .TextMatrix(.Rows - 1, IngredientCol(eGDIngredientCols_IngredientID)) = "-1"
                
                strText = "Click here to add a new ingredient"
                lStartCol = IngredientCol(eGDIngredientCols_Ingredient)
                lEndCol = IngredientCol(eGDIngredientCols_DryFeedPct)
                
                .Cell(flexcpText, .Rows - 1, lStartCol, .Rows - 1, lEndCol) = strText
                .Cell(flexcpForeColor, .Rows - 1, lStartCol, .Rows - 1, lEndCol) = vbBlue
                .Cell(flexcpFontUnderline, .Rows - 1, lStartCol, .Rows - 1, lEndCol) = True
                .Cell(flexcpAlignment, .Rows - 1, lStartCol, .Rows - 1, lEndCol) = flexAlignLeftCenter
                
                .Redraw = nRedraw
            End With
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleManage.AddClickHereLine"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowIsClickHereLine
'' Description: Determine if the given row in the grid is the click here line
'' Inputs:      Row
'' Returns:     True if click here line, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RowIsClickHereLine(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    Select Case m.nMode
        Case eGDManageMode_Ingredients
            bReturn = (fgObjects.TextMatrix(lRow, IngredientCol(eGDIngredientCols_IngredientID)) = "-1")
    
    End Select
    
    RowIsClickHereLine = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleManage.RowIsClickHereLine"
    
End Function

