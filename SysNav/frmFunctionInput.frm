VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmFunctionInput 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Maintenance"
   ClientHeight    =   4020
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VSFlex7LCtl.VSFlexGrid vsInputs 
      Height          =   3120
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   6195
      _cx             =   10927
      _cy             =   5503
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
   Begin HexUniControls.ctlUniButtonImageXP Corner 
      Height          =   285
      Left            =   6585
      TabIndex        =   2
      Top             =   3780
      Visible         =   0   'False
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
      Caption         =   "frmFunctionInput.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmFunctionInput.frx":002C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmFunctionInput.frx":004C
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   6465
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
      Caption         =   "frmFunctionInput.frx":0068
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmFunctionInput.frx":0096
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmFunctionInput.frx":00B6
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdSave 
      Height          =   375
      Left            =   6450
      TabIndex        =   0
      Top             =   150
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmFunctionInput.frx":00D2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmFunctionInput.frx":00FC
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmFunctionInput.frx":011C
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP txtPreview 
      Height          =   615
      Left            =   90
      Top             =   3285
      Width           =   6195
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
      Caption         =   "frmFunctionInput.frx":0138
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frmFunctionInput.frx":0158
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmFunctionInput.frx":0178
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmFunctionInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmFunctionInput.frm
'' Description: Allows the user to enter in information for a DLL Function Input
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    astrArgNames As cGdArray
    Lists As cLists
    InputTypes As cParmTypes
    Input As cInput
    bOK As Boolean
End Type
Private m As mPrivate

'Rows
Private Enum eGDRows
    eGDRow_InputName = 1
    eGDRow_InputDesc = 2
    eGDRow_Order = 3
    eGDRow_InputType = 4
    eGDRow_ListType = 5
    eGDRow_Required = 6
    eGDRow_DefaultValue = 7
    eGDRow_FillPre = 8
    eGDRow_FillPost = 9
    eGDRow_ValidFrom = 10
    eGDRow_ValidTo = 11
End Enum
Private Const kGridRows = 12

'Columns
Private Enum eGDCols
    eGDCol_ItemLabel = 0
    eGDCol_Value = 1
    eGDCol_Keys = 2
End Enum
Private Const kGridColumns = 3

Private Const bSave = 0
Private Const bCancel = 1

Private Function GDRow(ByVal lRow As eGDRows) As Long
    GDRow = lRow
End Function
Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .AllowBigSelection = False
        .AllowSelection = True
        .HighLight = flexHighlightWithFocus
        .TabBehavior = flexTabCells
        .Editable = True
        .ExtendLastCol = True
        .ExplorerBar = flexExNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        .ScrollTips = True
        .ScrollTrack = True
        .Ellipsis = flexEllipsisEnd
        .Cols = 3
        .Rows = kGridRows
        .FixedCols = 1
        .FixedRows = 1
        
        .TextMatrix(0, GDCol(eGDCol_ItemLabel)) = "Item"
        .TextMatrix(0, GDCol(eGDCol_Value)) = "Value"
        .TextMatrix(GDRow(eGDRow_InputName), GDCol(eGDCol_ItemLabel)) = "Input Name"
        .TextMatrix(GDRow(eGDRow_InputDesc), GDCol(eGDCol_ItemLabel)) = "Description"
        .TextMatrix(GDRow(eGDRow_Order), GDCol(eGDCol_ItemLabel)) = "Order"
        .TextMatrix(GDRow(eGDRow_InputType), GDCol(eGDCol_ItemLabel)) = "Input Type"
        .TextMatrix(GDRow(eGDRow_ListType), GDCol(eGDCol_ItemLabel)) = "List Type"
        .TextMatrix(GDRow(eGDRow_Required), GDCol(eGDCol_ItemLabel)) = "Required"
        .TextMatrix(GDRow(eGDRow_DefaultValue), GDCol(eGDCol_ItemLabel)) = "Default Value"
        .TextMatrix(GDRow(eGDRow_FillPre), GDCol(eGDCol_ItemLabel)) = "Fill words before"
        .TextMatrix(GDRow(eGDRow_FillPost), GDCol(eGDCol_ItemLabel)) = "Fill words after"
        .TextMatrix(GDRow(eGDRow_ValidFrom), GDCol(eGDCol_ItemLabel)) = "Valid From"
        .TextMatrix(GDRow(eGDRow_ValidTo), GDCol(eGDCol_ItemLabel)) = "Valid To"
        
        ' Make the Required field a check box
        .Cell(flexcpChecked, GDRow(eGDRow_Required), GDCol(eGDCol_Value)) = flexChecked
        
        'Out dated fields
        .RowHidden(GDRow(eGDRow_ValidFrom)) = True
        .RowHidden(GDRow(eGDRow_ValidTo)) = True
        .RowHidden(GDRow(eGDRow_FillPre)) = True
        .RowHidden(GDRow(eGDRow_FillPost)) = True
        .RowHidden(GDRow(eGDRow_ListType)) = True
        .RowHidden(GDRow(eGDRow_Order)) = True
        
        .ColHidden(GDCol(eGDCol_Keys)) = True
        .ColAlignment(GDCol(eGDCol_Value)) = flexAlignLeftCenter
        
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        .AutoSize 0, .Cols - 1
        SetColumnWidths
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmFunctionInput.InitGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetColumnWidths
'' Description: Make sure that column widths do not exceed max limits
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetColumnWidths()
On Error GoTo ErrSection:
    
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        If .ColWidth(GDCol(eGDCol_ItemLabel)) > 1500 Then .ColWidth(GDCol(eGDCol_ItemLabel)) = 1500
        If .ColWidth(GDCol(eGDCol_Value)) > 2500 Then .ColWidth(GDCol(eGDCol_Value)) = 2500
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionInput.SetColumnWidths", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load up the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With vsInputs
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .TextMatrix(GDRow(eGDRow_InputName), GDCol(eGDCol_Value)) = m.Input.ParmName
        .TextMatrix(GDRow(eGDRow_InputDesc), GDCol(eGDCol_Value)) = m.Input.ParmDesc
        .TextMatrix(GDRow(eGDRow_Order), GDCol(eGDCol_Value)) = m.Input.ParmSeq
        If m.InputTypes.Found(Str(m.Input.ParmTypeID)) Then
            .TextMatrix(GDRow(eGDRow_InputType), GDCol(eGDCol_Value)) = m.InputTypes.Item(CStr(m.Input.ParmTypeID)).ParmType
            .TextMatrix(GDRow(eGDRow_InputType), GDCol(eGDCol_Keys)) = m.Input.ParmTypeID
        End If
        If m.Lists.Found(CStr(m.Input.ListID)) Then
            .TextMatrix(GDRow(eGDRow_ListType), GDCol(eGDCol_Value)) = m.Lists.Item(CStr(m.Input.ListID)).ListName
            .TextMatrix(GDRow(eGDRow_ListType), GDCol(eGDCol_Keys)) = m.Input.ListID
        End If
        If m.Input.Required Then
            .Cell(flexcpChecked, GDRow(eGDRow_Required), GDCol(eGDCol_Value)) = flexChecked
        Else
            .Cell(flexcpChecked, GDRow(eGDRow_Required), GDCol(eGDCol_Value)) = flexUnchecked
        End If
        .TextMatrix(GDRow(eGDRow_DefaultValue), GDCol(eGDCol_Value)) = m.Input.DefaultValue
        .TextMatrix(GDRow(eGDRow_FillPre), GDCol(eGDCol_Value)) = m.Input.FillPre
        .TextMatrix(GDRow(eGDRow_FillPost), GDCol(eGDCol_Value)) = m.Input.FillPost
        .TextMatrix(GDRow(eGDRow_ValidFrom), GDCol(eGDCol_Value)) = m.Input.FromValue
        .TextMatrix(GDRow(eGDRow_ValidTo), GDCol(eGDCol_Value)) = m.Input.ToValue
        .AutoSize 0, .Cols - 1
        SetColumnWidths
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionInput.LoadGrid", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Unload the form without saving the input
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionInput.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSave_Click
'' Description: Save the input the user is working on
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSave_Click()
On Error GoTo ErrSection:
    
    Dim xInput As cInput                ' Temporary Input for saving
    
    Screen.MousePointer = vbHourglass
    
    With vsInputs
        If .TextMatrix(GDRow(eGDRow_InputType), GDCol(eGDCol_Keys)) = "4" Then
            If .TextMatrix(GDRow(eGDRow_DefaultValue), GDCol(eGDCol_Value)) = "" Then
                .TextMatrix(GDRow(eGDRow_DefaultValue), GDCol(eGDCol_Value)) = "Close"
            End If
        End If
    End With
    
    Set xInput = New cInput
    With xInput
        .DefaultValue = vsInputs.TextMatrix(GDRow(eGDRow_DefaultValue), GDCol(eGDCol_Value))
        .FillPost = vsInputs.TextMatrix(GDRow(eGDRow_FillPost), GDCol(eGDCol_Value))
        .FillPre = vsInputs.TextMatrix(GDRow(eGDRow_FillPre), GDCol(eGDCol_Value))
        .FromValue = vsInputs.Cell(flexcpValue, GDRow(eGDRow_ValidFrom), GDCol(eGDCol_Value))
        .ToValue = vsInputs.Cell(flexcpValue, GDRow(eGDRow_ValidTo), GDCol(eGDCol_Value))
        .ListID = CLng(ValOfText(vsInputs.TextMatrix(GDRow(eGDRow_ListType), GDCol(eGDCol_Keys))))
        .ParmDesc = vsInputs.TextMatrix(GDRow(eGDRow_InputDesc), GDCol(eGDCol_Value))
        .ParmName = vsInputs.TextMatrix(GDRow(eGDRow_InputName), GDCol(eGDCol_Value))
        .ParmSeq = vsInputs.Cell(flexcpValue, GDRow(eGDRow_Order), GDCol(eGDCol_Value))
        .ParmTypeID = CLng(ValOfText(vsInputs.TextMatrix(GDRow(eGDRow_InputType), GDCol(eGDCol_Keys))))
        .Required = (vsInputs.Cell(flexcpChecked, GDRow(eGDRow_Required), GDCol(eGDCol_Value)) = flexChecked)
        .Value = vsInputs.TextMatrix(GDRow(eGDRow_DefaultValue), GDCol(eGDCol_Value))
        .ValidateFunctionInput
    End With
        
    Set m.Input = xInput
    m.bOK = True
    Me.Hide
    
ErrExit:
    Screen.MousePointer = vbDefault
    Set xInput = Nothing
    Exit Sub
    
ErrSection:
    Screen.MousePointer = vbDefault
    RaiseError "frmFunctionInput.cmdSave.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionInput.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Ask the user if they want to save changes before unloading
'' Inputs:      Whether or not to Cancel the unload, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    Dim strResponse As String           ' Response from the ask box
    
    If UnloadMode = vbFormControlMenu Then
        strResponse = InfBox("Do you want to save your changes?", "?", "+Yes|No|-Cancel", "Confirmation")
        Select Case UCase(strResponse)
            Case "C"
                Cancel = True
                Exit Sub
            Case "Y"
                cmdSave_Click
                Exit Sub
            Case "N"
                m.bOK = False
                Cancel = True
                Me.Hide
        End Select
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionInput.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up before unloading the form
'' Inputs:      Whether or not to Cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:
    
    Set m.InputTypes = Nothing
    Set m.Lists = Nothing
    Set m.Input = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionInput.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize stuff and size and center the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Me.Icon = Picture16(ToolbarIcon("ID_Functions"), , True)
    ReSizeMDIChildForm Me, Corner
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    Set m.InputTypes = New cParmTypes
    m.InputTypes.Load
    
    Set m.Lists = New cLists
    m.Lists.Load
    
    cmdSave.Enabled = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionInput.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_AfterEdit
'' Description: Validate stuff after the edit
'' Inputs:      Row and Column of the cell edited
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    Dim ID As Long
    
    Select Case Row
        Case GDRow(eGDRow_InputType)
            With vsInputs
                ID = CLng(ValOfText(.ComboData))
                If m.InputTypes.Found(ID) Then
                    .TextMatrix(GDRow(eGDRow_InputType), GDCol(eGDCol_Keys)) = CStr(ID)
                    If ID = 5 Then
                        .TextMatrix(GDRow(eGDRow_InputName), GDCol(eGDCol_Value)) = "Market1"
                        .TextMatrix(GDRow(eGDRow_DefaultValue), GDCol(eGDCol_Value)) = "Market1"
                        .Cell(flexcpChecked, GDRow(eGDRow_Required), GDCol(eGDCol_Value)) = flexChecked
                    ElseIf ID = 4 Then
                        If .TextMatrix(GDRow(eGDRow_DefaultValue), GDCol(eGDCol_Value)) = "" Then
                            .TextMatrix(GDRow(eGDRow_DefaultValue), GDCol(eGDCol_Value)) = "Close"
                        End If
                    End If
                End If
            End With
        Case GDRow(eGDRow_ListType)
            With vsInputs
                ID = CLng(ValOfText(.ComboData))
                If m.Lists.Found(ID) Then
                    .TextMatrix(GDRow(eGDRow_ListType), GDCol(eGDCol_Value)) = m.Lists.Item(CStr(ID)).ListName
                    .TextMatrix(GDRow(eGDRow_ListType), GDCol(eGDCol_Keys)) = CStr(ID)
                End If
            End With
        Case GDRow(eGDRow_InputName), GDRow(eGDRow_DefaultValue)
            With vsInputs
                If CLng(ValOfText(.TextMatrix(GDRow(eGDRow_InputType), GDCol(eGDCol_Keys)))) = 5 Then
                    .TextMatrix(GDRow(eGDRow_InputName), GDCol(eGDCol_Value)) = "Market1"
                    .TextMatrix(GDRow(eGDRow_DefaultValue), GDCol(eGDCol_Value)) = "Market1"
                End If
            End With
        Case GDRow(eGDRow_Required)
            With vsInputs
                If CLng(ValOfText(.TextMatrix(GDRow(eGDRow_InputType), GDCol(eGDCol_Keys)))) = 5 Then
                    .Cell(flexcpChecked, GDRow(eGDRow_Required), GDCol(eGDCol_Value)) = flexChecked
                End If
            End With
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionInput.vsInputs.AfterEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_AfterRowColChange
'' Description: On a row/column change, try to set the cell to edit mode
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    With vsInputs
        If NewRow <> GDRow(eGDRow_Required) Then .EditCell
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionInput.vsInputs.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_BeforeEdit
'' Description: Set up the combo boxes if necessary
'' Inputs:      Row and Column of cell edited, Whether or not to Cancel edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim s As String
    Dim X As Integer
        
    vsInputs.ComboList = ""
    If vsInputs.Row = GDRow(eGDRow_InputType) Then
        For X = 1 To m.InputTypes.Count
            Select Case m.InputTypes.Item(X).ParmTypeID
                Case 1:
                    s = s & "|#" & m.InputTypes.Item(X).ParmTypeID & ";" & _
                        "Single Number (Double)"
                Case 2:
                    s = s & "|#" & m.InputTypes.Item(X).ParmTypeID & ";" & _
                        "Text (String)"
                Case 3:
                    s = s & "|#" & m.InputTypes.Item(X).ParmTypeID & ";" & _
                        "Series of True/False (cGdArray of 0's and 1's)"
                Case 4:
                    s = s & "|#" & m.InputTypes.Item(X).ParmTypeID & ";" & _
                        "Series of Numbers (cGdArray of Doubles)"
                Case 5:
                    s = s & "|#" & m.InputTypes.Item(X).ParmTypeID & ";" & _
                        "Market (cGdBars)"
                Case 6:
                    s = s & "|#" & m.InputTypes.Item(X).ParmTypeID & ";" & _
                        "Single True/False (0 or 1)"
            End Select
        Next X
        If Left(s, 1) = "|" Then s = Right(s, Len(s) - 1)
        vsInputs.ComboList = s
    End If
    
    If vsInputs.Row = GDRow(eGDRow_ListType) Then
        If vsInputs.TextMatrix(GDRow(eGDRow_InputType), 1) = "Text" Then
            For X = 1 To m.Lists.Count
                s = s & "|#" & m.Lists.Item(X).ListID & ";" & _
                    m.Lists.Item(X).ListName
            Next X
            vsInputs.ComboList = s
        End If
    End If
    cmdSave.Enabled = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionInput.vsInputs.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsInputs_ValidateEdit
'' Description: Validate what the user entered
'' Inputs:      Row and Column of cell edited, Whether or not to Cancel edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsInputs_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Select Case Row
        Case GDRow(eGDRow_InputName)
            With vsInputs
                If UCase(.EditText) <> UCase(m.Input.ParmName) Then
                    If m.astrArgNames.BinarySearch(.EditText, , eGdSort_IgnoreCase) = True Then
                        InfBox "h=Error ; i=! ; " & .EditText & " already exists.  Please rename argument"
                        Cancel = True
                    End If
                End If
            End With
        
        Case GDRow(eGDRow_InputType)
            With vsInputs
                If .ComboData = "5" And .TextMatrix(GDRow(eGDRow_InputName), GDCol(eGDCol_Value)) <> "Market1" Then
                    If m.astrArgNames.BinarySearch("Market1", , eGdSort_IgnoreCase) = True Then
                        InfBox "h=Error ; i=! ; Only one Market type argument allowed"
                        Cancel = True
                    End If
                End If
            End With
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionInput.vsInputs.ValidateEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Set up and show the form
'' Inputs:      Input to load and return, Argument Names
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(pInput As cInput, astrArgNames As cGdArray) As Boolean
On Error GoTo ErrSection:

    Set m.Input = pInput
    Set m.astrArgNames = astrArgNames
    m.astrArgNames.Sort
    
    vsInputs.Redraw = flexRDNone
    InitGrid
    LoadGrid
    vsInputs.Redraw = flexRDBuffered
    
    ShowForm Me, True
    
    If m.bOK = True Then Set pInput = m.Input
    
    ShowMe = m.bOK
    Unload Me

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmFunctionInput.ShowMe", eGDRaiseError_Raise

End Function

