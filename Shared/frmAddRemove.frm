VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.0#0"; "HexUniControls42.ocx"
Begin VB.Form frmAddRemove 
   Caption         =   "Arrange Grid"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8055
   Icon            =   "frmAddRemove.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4920
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLeft 
      Height          =   195
      Left            =   4680
      Picture         =   "frmAddRemove.frx":0442
      ScaleHeight     =   135
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   135
   End
   Begin HexUniControls.ctlUniFrameWL fraUpDown 
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   4080
      Width           =   1935
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
      Caption         =   "frmAddRemove.frx":048E
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAddRemove.frx":04AE
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAddRemove.frx":04CE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDown 
         Height          =   315
         Left            =   900
         TabIndex        =   12
         Top             =   45
         Width           =   1020
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
         Caption         =   "frmAddRemove.frx":04EA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAddRemove.frx":051E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAddRemove.frx":053E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdUp 
         Height          =   315
         Left            =   0
         TabIndex        =   11
         Top             =   45
         Width           =   900
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
         Caption         =   "frmAddRemove.frx":055A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAddRemove.frx":058A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAddRemove.frx":05AA
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraMiddle 
      Height          =   3135
      Left            =   3420
      TabIndex        =   2
      Top             =   420
      Width           =   1155
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
      Caption         =   "frmAddRemove.frx":05C6
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAddRemove.frx":05E6
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAddRemove.frx":0606
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDefaults 
         Height          =   510
         Left            =   120
         TabIndex        =   7
         Top             =   2580
         Width           =   900
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
         Caption         =   "frmAddRemove.frx":0622
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAddRemove.frx":0662
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAddRemove.frx":0682
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
         Height          =   510
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   900
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
         Caption         =   "frmAddRemove.frx":069E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAddRemove.frx":06CC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAddRemove.frx":06EC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRemove 
         Height          =   510
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   900
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
         Caption         =   "frmAddRemove.frx":0708
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAddRemove.frx":073C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAddRemove.frx":075C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   405
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   900
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
         Caption         =   "frmAddRemove.frx":0778
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAddRemove.frx":07A6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAddRemove.frx":07C6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   405
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   900
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
         Caption         =   "frmAddRemove.frx":07E2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAddRemove.frx":0808
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAddRemove.frx":0828
         RightToLeft     =   0   'False
      End
   End
   Begin VB.PictureBox picRight 
      Height          =   195
      Left            =   6502
      Picture         =   "frmAddRemove.frx":0844
      ScaleHeight     =   135
      ScaleWidth      =   75
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VSFlex7LCtl.VSFlexGrid fgUsed 
      Height          =   3675
      Left            =   4620
      TabIndex        =   9
      Top             =   360
      Width           =   3300
      _cx             =   5821
      _cy             =   6482
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
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
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
      OleDropMode     =   1
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7LCtl.VSFlexGrid fgAvailable 
      Height          =   3675
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3285
      _cx             =   5794
      _cy             =   6482
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
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
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
      OleDropMode     =   1
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin HexUniControls.ctlUniLabelXP lblUsed 
      Height          =   315
      Left            =   4560
      Top             =   120
      Width           =   3300
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
      Caption         =   "frmAddRemove.frx":0890
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmAddRemove.frx":08D6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAddRemove.frx":08F6
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblAvailable 
      Height          =   225
      Left            =   135
      Top             =   120
      Width           =   3285
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
      Caption         =   "frmAddRemove.frx":0912
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmAddRemove.frx":0956
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAddRemove.frx":0976
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmAddRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmAddRemove.frm
'' Description: Generic Add/Remove form to provide the user with two lists -
''              one list of available things to add and another list of things
''              already added.
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
'' ??/??/??  R Johnson   Created
'' 11/10/00  D Jarmuth   Added comments/formatting, fixed code, added ShowMe
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Enum eOrderMode
    eOrderMode_Ordered = 0
    eOrderMode_NotOrdered = 1
    eOrderMode_Alphabetical = 2
End Enum

Private Type mPrivate
    bOK As Boolean                ' True if OK, False if Cancelled
    Order As eOrderMode           ' True if order specific, False otherwise
    strDragSource As String       ' Source of the drag
    sDraggedItems As Variant      ' Items that are being dragged
    iDraggedRows As Variant       ' Rows that are being dragged
    aDefaultList As cGdArray
    aAll As cGdArray
End Type
Private m As mPrivate

Private Sub cmdDefaults_Click()
On Error GoTo ErrSection:

    Dim i&

    With fgUsed
        .Redraw = flexRDNone
        .Rows = .FixedRows
        For i = 0 To m.aDefaultList.Size - 1
            .AddItem m.aDefaultList(i)
        Next
        If m.Order = eOrderMode_Alphabetical Then
            .Select .FixedRows, 0
            .Sort = flexSortStringAscending
        End If
        .Redraw = flexRDBuffered
    End With
    
    FillAvailableGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.cmdDefaults.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemove_Click
'' Description: Moves an item from the "used" side to the "available" side
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemove_Click()
On Error GoTo ErrSection:

    Dim lRows As Long                   ' Index into a for loop
    Dim iCols As Integer                ' Index into a for loop
    Dim strData As String               ' Data from the row of the grid
    Dim alSelectedRows() As Long        ' Selected rows in the grid
    Dim lNumSelected As Long            ' Number of rows that are selected
    
    With fgUsed
        ' Set up the array of selected rows in the used grid
        lNumSelected = .SelectedRows
        ReDim alSelectedRows(lNumSelected) As Long
        For lRows = 0 To lNumSelected - 1
            alSelectedRows(lRows) = .SelectedRow(lRows)
        Next lRows
        
        ' Walk through the selected rows and move them to the available grid
        For lRows = lNumSelected - 1 To 0 Step -1
            ' Build the data string for the given row
            strData = .TextMatrix(alSelectedRows(lRows), 0)
            For iCols = 1 To .Cols - 1
                strData = strData & vbTab & .TextMatrix(alSelectedRows(lRows), iCols)
            Next iCols
            
            ' Add the row to the available grid and remove from the used grid
            'fgAvailable.AddItem strData
            .RemoveItem alSelectedRows(lRows)
            'fgAvailable.Select fgAvailable.Rows - 1, 0
        Next lRows
    End With
    
    FillAvailableGrid
    EnableButtons
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.cmdRemove.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user clicks on the cancel button, make the ok false and
''              unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:
    
    m.bOK = False
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDown_Click
'' Description: If the user clicks on the down button, move an item in the
''              "used" list down one position in the list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDown_Click()
On Error GoTo ErrSection:
    
    Dim strTemp1 As String              ' Temporary string for swapping
    Dim strTemp2 As String              ' Temporary string for swapping
  
    ' Make sure we can still move down, then complete the swap
    If (fgUsed.Row + 1) < fgUsed.Rows Then
        strTemp1 = fgUsed.Text
        fgUsed.Row = fgUsed.Row + 1
        strTemp2 = fgUsed.Text
        fgUsed.Text = strTemp1
        fgUsed.Row = fgUsed.Row - 1
        fgUsed.Text = strTemp2
        fgUsed.Row = fgUsed.Row + 1
        fgUsed.ShowCell fgUsed.Row, fgUsed.Col
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.cmdDown.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user clicks on the OK button, set ok to true and unload
''              the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    ' Need to do this for "Default" buttons, since hitting 'Enter'
    ' from another control does not trigger it's 'LostFocus' event.
    MoveFocus cmdOK

    m.bOK = True
    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAdd_Click
'' Description: If the user clicks on the add button, move the selected item
''              in the "available" list over to the "used" list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAdd_Click()
On Error GoTo ErrSection:
    
    Dim lRows As Long                   ' Index into a for loop
    Dim iCols As Integer                ' Index into a for loop
    Dim strData As String               ' Data from the row of the grid
    Dim alSelectedRows() As Long        ' Selected rows in the grid
    Dim lNumSelected As Long            ' Number of rows that are selected
    
    With fgAvailable
        ' Set up the array of selected rows in the used grid
        lNumSelected = .SelectedRows
        ReDim alSelectedRows(lNumSelected) As Long
        For lRows = 0 To lNumSelected - 1
            alSelectedRows(lRows) = .SelectedRow(lRows)
        Next lRows
        
        ' Walk through the selected rows and move them to the available grid
        For lRows = lNumSelected - 1 To 0 Step -1
            ' Build the data string for the given row
            strData = .TextMatrix(alSelectedRows(lRows), 0)
            For iCols = 1 To .Cols - 1
                strData = strData & vbTab & .TextMatrix(alSelectedRows(lRows), iCols)
            Next iCols
            
            ' Add the row to the available grid and remove from the used grid
            fgUsed.AddItem strData
            '.RemoveItem alSelectedRows(lRows)
            fgUsed.Select fgUsed.Rows - 1, 0
        Next lRows
    End With
    
    FillAvailableGrid
    MoveFocus fgUsed
    EnableButtons
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.cmdAdd.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdUp_Click
'' Description: If the user clicks on the Up button, move an item in the "used"
''              list up one position in the list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdUp_Click()
On Error GoTo ErrSection:
    
    Dim strTemp1 As String              ' Temporary string for swapping
    Dim strTemp2 As String              ' Temporary string for swapping
  
    ' Make sure we aren't already at the top of the list, otherwise swap rows
    If (fgUsed.Row - 1) >= 0 Then
        strTemp1 = fgUsed.Text
        fgUsed.Row = fgUsed.Row - 1
        strTemp2 = fgUsed.Text
        fgUsed.Text = strTemp1
        fgUsed.Row = fgUsed.Row + 1
        fgUsed.Text = strTemp2
        fgUsed.Row = fgUsed.Row - 1
        fgUsed.ShowCell fgUsed.Row, fgUsed.Col
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.cmdUp.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAvailable_BeforeMouseDown
'' Description: Used to initiate the drag operation for the fgAvailable grid
'' Inputs:      Mouse button pressed, Shift status, X Location, Y Location,
''              Whether or not to cancel the operation
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAvailable_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Increment variable
    Dim lMouseRow As Long               ' Current mouse row
    
    Static slRow As Long                ' Last row that was selected
    Static slXXX As Long

    With fgAvailable
    
        If .MouseRow < .FixedRows Then Exit Sub
        
        ' Capture the mouse row in case this takes a while...
        lMouseRow = .MouseRow

        ' The Shift key is being pressed
        If Shift And vbShiftMask Then
            slXXX = 0
            
            ' If the Control key is not down, clear the current selection and
            ' start over
            If (Shift And vbCtrlMask) = 0 Then
                .Row = lMouseRow
            End If
            
            ' Select everything in between the last row and the current mouse row
            If slRow < lMouseRow Then
                For lIndex = slRow To lMouseRow
                    .IsSelected(lIndex) = True
                Next lIndex
            ElseIf slRow > lMouseRow Then
                For lIndex = slRow To lMouseRow Step -1
                    .IsSelected(lIndex) = True
                Next lIndex
            Else
                .IsSelected(lMouseRow) = True
            End If
            
        ' The Control key is being pressed, but not the Shift key
        ElseIf Shift And vbCtrlMask Then
            slXXX = 0
            
            ' Toggle the selection of the row being clicked on
            .IsSelected(lMouseRow) = Not .IsSelected(lMouseRow)
        
        ' No key is being pressed (that we care about)
        Else
            ' If the current row is not selected or it has been clicked twice in
            ' a row, then clear out the current selection and start over
            If .IsSelected(lMouseRow) = False Or lMouseRow = slXXX Then
                .Row = lMouseRow
                slXXX = 0
            Else
                slXXX = lMouseRow
            End If
            
            .IsSelected(lMouseRow) = True
        End If
        
        ' If the Shift key was not being pressed, then change the last saved row
        If (Shift And vbShiftMask) = 0 Then
            If .SelectedRows > 0 Then
                slRow = lMouseRow
            Else
                slRow = 0&
            End If
        End If
        
        ' Use OLEDrag method to start manual OLE drag operation
        ' this will fire the OLEStartDrag event, which we will use
        ' to fill the DataObject with the data we want to drag.
        .OLEDrag

        ' Tell grid control to ignore mouse movements until the
        ' mouse button goes up again
        Cancel = True
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.fgAvailable.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgAvailable_DblClick()
On Error GoTo ErrSection:

    cmdAdd_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.fgAvailable.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAvailable_OLEDragDrop
'' Description: Drop dragged items into fgAvailable grid
'' Inputs:      Data to drop, Effect, Mouse button pressed, Shift status, X
''              Location, Y Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAvailable_OLEDragDrop(Data As VSFlex7LCtl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error GoTo ErrSection:

    Dim iIndex As Integer               ' Increment variable

    ' We only accept drags from the fgUsed grid
    If m.strDragSource = "list" Then
        ' Build a list of the selected items
        For iIndex = 0 To UBound(m.iDraggedRows)
            fgUsed.Row = m.iDraggedRows(iIndex)
            m.sDraggedItems(iIndex) = fgUsed.Text
        Next iIndex
     
        ' Remove them from the old list
        For iIndex = 0 To UBound(m.sDraggedItems) - 1
            fgUsed.RemoveItem m.iDraggedRows(iIndex) - iIndex
            If fgUsed.Row > m.iDraggedRows(iIndex) Then fgUsed.Row = fgUsed.Row - 1
        Next iIndex
     
        ' Select new fgUsed row
        If fgUsed.Row <> -1 Then fgUsed.Select fgUsed.Row, 0
     
        ' Add to the new list
        'For iIndex = 0 To UBound(m.sDraggedItems) - 1
            'fgAvailable.AddItem m.sDraggedItems(iIndex)
        'Next iIndex
        FillAvailableGrid
        
    End If
    
    MoveFocus fgAvailable
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.fgAvailable.OLEDragDrop", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgAvailable_OLEStartDrag
'' Description: Begin Drag Procedure
'' Inputs:      Data to drag, Allowed effects of the drag
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgAvailable_OLEStartDrag(Data As VSFlex7LCtl.VSDataObject, AllowedEffects As Long)
On Error GoTo ErrSection:

    Dim iIndex As Integer               ' Increment variable

    m.strDragSource = "items"
    fgAvailable.OLEDropMode = flexOLEDropManual
  
    ' Check for number of items to move
    If fgAvailable.SelectedRows > 1 Then
        ReDim m.sDraggedItems(fgAvailable.SelectedRows)
        ReDim m.iDraggedRows(fgAvailable.SelectedRows)
   
        ' Store the row information
        For iIndex = 0 To fgAvailable.SelectedRows - 1
            m.iDraggedRows(iIndex) = fgAvailable.SelectedRow(iIndex)
        Next iIndex
    Else
        ReDim m.sDraggedItems(1)
        ReDim m.iDraggedRows(1)
        m.iDraggedRows(0) = fgAvailable.Row
    End If
    
    ' Set contents of data object for manual drag
    ' (Put this code in so that the program would not crash when the mouse is
    ' moved off the form - 11/10/00 DAJ)
    Dim s$
    s = fgAvailable.Clip
    Data.SetData s, vbCFText
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.fgAvailable.OLEStartDrag", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgUsed_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    EnableButtons

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.fgUsed.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgUsed_BeforeMouseDown
'' Description: Initiate the drag procedure
'' Inputs:      Mouse button pressed, Shift status, X Location, Y Location,
''              Whether or not to cancel the operation
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgUsed_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Increment variable
    Dim lMouseRow As Long               ' Current mouse row
    Dim lMouseCol As Long               ' Current mouse column
    
    Static slRow As Long                ' Last row that was selected
    Static slXXX As Long

    With fgUsed
    
        If .MouseRow < .FixedRows Then Exit Sub
        
        ' Capture the mouse row in case this takes a while...
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol

        ' The Shift key is being pressed
        If Shift And vbShiftMask Then
            slXXX = 0
            
            ' If the Control key is not down, clear the current selection and
            ' start over
            If (Shift And vbCtrlMask) = 0 Then
                .Row = lMouseRow
            End If
            
            ' Select everything in between the last row and the current mouse row
            If slRow < lMouseRow Then
                For lIndex = slRow To lMouseRow
                    .IsSelected(lIndex) = True
                Next lIndex
            ElseIf slRow > lMouseRow Then
                For lIndex = slRow To lMouseRow Step -1
                    .IsSelected(lIndex) = True
                Next lIndex
            Else
                .IsSelected(lMouseRow) = True
            End If
            
        ' The Control key is being pressed, but not the Shift key
        ElseIf Shift And vbCtrlMask Then
            slXXX = 0
            
            ' Toggle the selection of the row being clicked on
            .IsSelected(lMouseRow) = Not .IsSelected(lMouseRow)
        
        ' No key is being pressed (that we care about)
        Else
            ' If the current row is not selected or it has been clicked twice in
            ' a row, then clear out the current selection and start over
            If .IsSelected(lMouseRow) = False Or lMouseRow = slXXX Then
                .Row = lMouseRow
                slXXX = 0
            Else
                slXXX = lMouseRow
            End If
            
            .IsSelected(lMouseRow) = True
        End If
        
        ' If the Shift key was not being pressed, then change the last saved row
        If (Shift And vbShiftMask) = 0 Then
            If .SelectedRows > 0 Then
                slRow = lMouseRow
            Else
                slRow = 0&
            End If
        End If
        
        ' Use OLEDrag method to start manual OLE drag operation
        ' this will fire the OLEStartDrag event, which we will use
        ' to fill the DataObject with the data we want to drag.
        .OLEDrag

        ' Tell grid control to ignore mouse movements until the
        ' mouse button goes up again
        Cancel = True
    
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.fgUsed.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgUsed_DblClick()
On Error GoTo ErrSection:

    cmdRemove_Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.fgUsed.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgUsed_OLEDragDrop
'' Description: Drop information into fgUsed
'' Inputs:      Data to drop, Effects, Mouse button pressed, Shift status, X
''              Location, Y Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgUsed_OLEDragDrop(Data As VSFlex7LCtl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error GoTo ErrSection:

    Dim iIndexValue As Integer          ' Row to drop information into
    Dim iIndex As Integer               ' Increment variable
    Dim bInSelect As Boolean            ' Boolean for checking mouse position

    bInSelect = False
    If m.Order = eOrderMode_Ordered Then
        iIndexValue = fgUsed.MouseRow
    Else
        iIndexValue = -1
    End If

    ' Store the values and remove rows
    If m.strDragSource = "list" Then
        ' Store dragged item information
        For iIndex = 0 To UBound(m.iDraggedRows) - 1
            fgUsed.Row = m.iDraggedRows(iIndex)
            m.sDraggedItems(iIndex) = fgUsed.Text
            If iIndexValue = m.iDraggedRows(iIndex) Then bInSelect = True
            If bInSelect = True Then GoTo ExitSub
        Next iIndex
  
        ' Remove dragged rows
        For iIndex = 0 To UBound(m.sDraggedItems) - 1
            fgUsed.RemoveItem m.iDraggedRows(iIndex) - iIndex
            If m.iDraggedRows(iIndex) < iIndexValue Then iIndexValue = iIndexValue - 1
        Next iIndex
    Else
        ' Store dragged item information
        For iIndex = 0 To UBound(m.iDraggedRows) - 1
            fgAvailable.Row = m.iDraggedRows(iIndex)
            m.sDraggedItems(iIndex) = fgAvailable.Text
        Next iIndex
        ' Remove dragged rows
        'For iIndex = 0 To UBound(m.sDraggedItems) - 1
        '    fgAvailable.RemoveItem m.iDraggedRows(iIndex) - iIndex
        'Next iIndex
    End If

    ' If we don't have a valid column add to the end of the control
    If iIndexValue <= -1 Then
        ' When doing an additem without a row value do them in incrementing order
        For iIndex = 0 To UBound(m.sDraggedItems) - 1
            fgUsed.AddItem m.sDraggedItems(iIndex)
            fgUsed.IsSelected(fgUsed.Rows - 1) = True
            If iIndex = 0 Then fgUsed.Row = fgUsed.Rows - 1
        Next iIndex
        fgUsed.ShowCell fgUsed.Rows - 1, 0
    Else
        fgUsed.Row = iIndexValue
        fgUsed.IsSelected(fgUsed.Row) = False
        
        ' When doing an additem with a row value we need to do last item first
        For iIndex = UBound(m.sDraggedItems) - 1 To 0 Step -1
            fgUsed.AddItem m.sDraggedItems(iIndex), fgUsed.Row
            fgUsed.IsSelected(fgUsed.Row) = True
        Next iIndex
        fgUsed.ShowCell fgUsed.Row, 0
    End If
    
    If m.Order = eOrderMode_Alphabetical Then
        fgUsed.Select fgUsed.FixedRows, 0
        fgUsed.Sort = flexSortStringAscending
    End If
    
    FillAvailableGrid
    
ExitSub:
    picLeft.Visible = False
    picRight.Visible = False
    fgAvailable.OLEDropMode = flexOLEDropNone
    
    MoveFocus fgUsed
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.fgUsed.OLEDragDrop", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgUsed_OLEDragOver
'' Description: What to do while dragging over the fgUsed list
'' Inputs:      Data being dragged, Effects, Mouse button pressed, Shift status
''              X Location, Y Location, Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgUsed_OLEDragOver(Data As VSFlex7LCtl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, State As Integer)
On Error GoTo ErrSection:

    If m.Order <> eOrderMode_Ordered Then Exit Sub
    
    ' Initialize pictures for drag routine
    If picLeft.Visible = False And fgUsed.Rows > 0 Then
        picLeft.Left = fgUsed.Left - picLeft.Width
        picRight.Left = fgUsed.Left + fgUsed.Width
        
        picLeft.Visible = True
        picRight.Visible = True
    End If

    ' While dragging move the arrow pictures
    If picLeft.Visible = True Then
        'If fgUsed.MouseRow = -1 Then
        If fgUsed.MouseRow < fgUsed.FixedRows Then
            picLeft.Top = (fgUsed.Rows - fgUsed.TopRow + fgUsed.FixedRows) * fgUsed.RowHeight(0) + fgUsed.Top - (picLeft.Height / 2)
            picRight.Top = picLeft.Top
        Else
            picLeft.Top = (fgUsed.MouseRow - fgUsed.TopRow + fgUsed.FixedRows) * fgUsed.RowHeight(0) + fgUsed.Top - (picLeft.Height / 2)
            picRight.Top = picLeft.Top
        End If
    End If

    ' If leaving control hide arrow pictures
    If State = vbLeave Then
        picLeft.Visible = False
        picRight.Visible = False
        Exit Sub
    End If

    ' Scroll up
    If fgUsed.MouseRow = fgUsed.TopRow And fgUsed.TopRow <> 0 Then
        fgUsed.TopRow = fgUsed.TopRow - 1
        Exit Sub
    End If

    ' Scroll down
    If fgUsed.Rows > 0 Then
        If fgUsed.MouseRow = -1 And Y > fgUsed.Top + fgUsed.Height - fgUsed.RowHeight(0) Then
            fgUsed.TopRow = fgUsed.TopRow + 1
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.fgUsed.OLEDragOver", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgUsed_OLEStartDrag
'' Description: Start drag from fgUsed
'' Inputs:      Data to drag, Effects allowed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgUsed_OLEStartDrag(Data As VSFlex7LCtl.VSDataObject, AllowedEffects As Long)
On Error GoTo ErrSection:

    Dim iIndex As Integer               ' Increment variable
  
    m.strDragSource = "list"
    fgAvailable.OLEDropMode = flexOLEDropManual
  
    ' Retrieve and store which rows were selected
    If fgUsed.SelectedRows > 1 Then
        ReDim m.sDraggedItems(fgUsed.SelectedRows)
        ReDim m.iDraggedRows(fgUsed.SelectedRows)
        
        ' Store the rows
        For iIndex = 0 To fgUsed.SelectedRows - 1
            m.iDraggedRows(iIndex) = fgUsed.SelectedRow(iIndex)
        Next iIndex
    Else
        ReDim m.sDraggedItems(1)
        ReDim m.iDraggedRows(1)
        m.iDraggedRows(0) = fgUsed.Row
        m.sDraggedItems(0) = fgUsed.Text
    End If
    
    ' Set contents of data object for manual drag
    ' (Put this code in so that the program would not crash when the mouse is
    ' moved off the form - 11/10/00 DAJ)
    Dim s$
    s = fgAvailable.Clip
    Data.SetData s, vbCFText
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.fgUsed.OLEStartDrag", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form is activated, reset the toolbars
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:
    
    If m.Order <> eOrderMode_Ordered Then
        fraUpDown.Visible = False
        Me.Height = fraUpDown.Top
    End If
    
    EnableButtons
    Screen.MousePointer = vbDefault
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.Form.Activate", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        If Not g.Help Is Nothing Then g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, initialize variables
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:
    
    ' Initialize form variables
    ReDim m.iDraggedRows(0)
    ReDim m.sDraggedItems(0)
    m.strDragSource = " "
    m.bOK = False
    
    'RH commented out fraMiddle.BorderStyle = vbBSNone
    'RH commented out fraUpDown.BorderStyle = vbBSNone
    
    g.Styler.StyleForm Me
    
    CenterTheForm Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Set minimum height and width for the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next
    
    Dim lUpDownTop&
    
    If LimitFormSize(Me, fraMiddle.Width + fraUpDown.Width * 2 + fgAvailable.Left, _
        fraMiddle.Top + fraMiddle.Height + fraUpDown.Height) Then Exit Sub
    
    '(more efficient to do .Move if more than one resize on an object)
    If fraUpDown.Visible = True Then
        lUpDownTop = Me.ScaleHeight - fraUpDown.Height - 30
    Else
        lUpDownTop = Me.ScaleHeight
    End If
    
    fraMiddle.Left = (Me.ScaleWidth - fraMiddle.Width) / 2
    lblUsed.Left = fraMiddle.Left + fraMiddle.Width
    
    With fgAvailable
        .Move .Left, .Top, fraMiddle.Left - .Left, lUpDownTop - .Top - 60
    End With
    With fgUsed
        .Move lblUsed.Left, .Top, _
                fgAvailable.Width, fgAvailable.Height
    End With
    
    fraUpDown.Move fgUsed.Left + ((fgUsed.Width - fraUpDown.Width) / 2), lUpDownTop
    
    Me.Refresh
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form is unloaded, do some final clean up
'' Inputs:      Whether or not to cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    'm.bOK = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadList
'' Description: Load the given control list with the given array data
'' Inputs:      Data to load, Control to load data into
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadList(List As cGdArray, cntl As Control)
On Error GoTo ErrSection:
    
    Dim iIndex As Integer               ' Increment variable
    
    For iIndex = 0 To List.Size - 1
        cntl.AddItem List.Item(iIndex)
    Next iIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.LoadList", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RetrieveList
'' Description: Retrieves the list of data from a given control
'' Inputs:      Control to receive data from
'' Returns:     Data in the control
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RetrieveList(cntl As Control) As cGdArray
On Error GoTo ErrSection:

    Dim List As New cGdArray
    Dim iIndex As Integer
  
    cntl.Col = 0
    For iIndex = 0 To cntl.Rows - 1
        cntl.Row = iIndex
        If cntl.Text <> "" Then List.Add cntl.Text
    Next iIndex
    Set RetrieveList = List
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAddRemove.RetrieveList", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Sets up the two list controls with the arrays that are passed
''              in and then when the user closes the form, it sets the two
''              arrays with the appropriate data from the list controls
'' Inputs:      Available List (can have duplicates of Used List), Used List
'' Returns:     True if OK was pressed, False if Cancel was pressed
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByRef AvailableList As cGdArray, ByRef UsedList As cGdArray, OrderSpecific As eOrderMode, _
        Optional ByVal DefaultList As cGdArray = Nothing, Optional ByVal strFormCaption$ = "Arrange Fields") As Boolean
On Error GoTo ErrSection:

    Dim i As Integer               ' Index into for loops
    
    Screen.MousePointer = vbHourglass
    
    ' build ALL list
    Set m.aAll = New cGdArray
    m.aAll.Create eGDARRAY_Strings
    For i = 0 To UsedList.Size - 1
        m.aAll.Add UsedList(i)
    Next
    For i = 0 To AvailableList.Size - 1
        m.aAll.Add AvailableList(i)
    Next
    m.aAll.Sort eGdSort_DeleteDuplicates Or eGdSort_IgnoreCase
    
    ' see if default list was passed
    If DefaultList Is Nothing Then
        Set m.aDefaultList = New cGdArray
    Else
        Set m.aDefaultList = DefaultList
    End If
    If m.aDefaultList.Size = 0 Then
        cmdDefaults.Visible = False
        fraMiddle.Height = cmdDefaults.Top - 120
    End If
        
    ' Fill in the used list control
    With fgUsed
        .Redraw = flexRDNone
        .Rows = .FixedRows
        For i = 0 To UsedList.Size - 1
            .AddItem UsedList(i)
        Next
        If OrderSpecific = eOrderMode_Alphabetical Then
            .Select .FixedRows, 0
            .Sort = flexSortStringAscending
        End If
        .Redraw = flexRDBuffered
    End With
    
    ' Fill in the available list control
    FillAvailableGrid
    
    fgAvailable.Row = -1
    fgUsed.Row = -1
    
    ' Show the form
    m.bOK = False
    m.Order = OrderSpecific
    Me.Caption = strFormCaption
    Show vbModal
    
    Set m.aDefaultList = Nothing
    Set m.aAll = Nothing
    
    ' If the user clicked on the OK button, return the new lists
    If m.bOK = True Then
        ' Fill in the available array
        AvailableList.Clear
        For i = fgAvailable.FixedRows To fgAvailable.Rows - 1
            AvailableList.Add fgAvailable.TextMatrix(i, 0)
        Next
        
        ' Fill in the used array
        UsedList.Clear
        For i = fgUsed.FixedRows To fgUsed.Rows - 1
            UsedList.Add fgUsed.TextMatrix(i, 0)
        Next
        Unload Me
    End If
    
    ShowMe = m.bOK
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAddRemove.ShowMe", eGDRaiseError_Raise

End Function

Private Sub EnableButtons()
On Error GoTo ErrSection:

    If fgUsed.Row >= fgUsed.FixedRows + 1 And fgUsed.Row < fgUsed.Rows Then
        Enable cmdUp
    Else
        Disable cmdUp
    End If
    If fgUsed.Row >= fgUsed.FixedRows And fgUsed.Row < fgUsed.Rows - 1 Then
        Enable cmdDown
    Else
        Disable cmdDown
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.EnableButtons", eGDRaiseError_Raise

End Sub

Private Sub FillAvailableGrid()
On Error GoTo ErrSection:

    Dim i&, aUsed As New cGdArray
    
    ' get list of what's in Used list
    For i = fgUsed.FixedRows To fgUsed.Rows - 1
        aUsed.Add fgUsed.TextMatrix(i, 0)
    Next
    aUsed.Sort eGdSort_IgnoreCase
    If fgUsed.Row >= fgUsed.FixedRows Then
        fgUsed.ShowCell fgUsed.Row, 0
    End If

    ' Fill in the available list control
    ' (ignore what's already used)
    With fgAvailable
        .Redraw = flexRDNone
        .Rows = .FixedRows
        For i = 0 To m.aAll.Size - 1
            If Not aUsed.BinarySearch(m.aAll(i), , eGdSort_IgnoreCase) Then
                .AddItem m.aAll(i)
            End If
        Next
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAddRemove.FillAvailableGrid", eGDRaiseError_Raise

End Sub


