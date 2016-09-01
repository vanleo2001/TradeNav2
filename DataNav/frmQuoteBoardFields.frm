VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmQuoteBoardFields 
   Caption         =   "Quote Board Fields"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraUpDown 
      Height          =   375
      Left            =   1020
      TabIndex        =   2
      Top             =   4260
      Width           =   2475
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
      Caption         =   "frmQuoteBoardFields.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmQuoteBoardFields.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQuoteBoardFields.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDown 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   0
         Width           =   1155
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
         Caption         =   "frmQuoteBoardFields.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQuoteBoardFields.frx":009C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQuoteBoardFields.frx":00BC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdUp 
         Height          =   315
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1155
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
         Caption         =   "frmQuoteBoardFields.frx":00D8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQuoteBoardFields.frx":0108
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQuoteBoardFields.frx":0128
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   2775
      Left            =   4020
      TabIndex        =   1
      Top             =   120
      Width           =   1095
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
      Caption         =   "frmQuoteBoardFields.frx":0144
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmQuoteBoardFields.frx":0170
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQuoteBoardFields.frx":0190
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdRemove 
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   1980
         Width           =   1095
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
         Caption         =   "frmQuoteBoardFields.frx":01AC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQuoteBoardFields.frx":01DA
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQuoteBoardFields.frx":01FA
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   1560
         Width           =   1095
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
         Caption         =   "frmQuoteBoardFields.frx":0216
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQuoteBoardFields.frx":0240
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQuoteBoardFields.frx":0260
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   1140
         Width           =   1095
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
         Caption         =   "frmQuoteBoardFields.frx":027C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQuoteBoardFields.frx":02A4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQuoteBoardFields.frx":02C4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   420
         Width           =   1095
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
         Caption         =   "frmQuoteBoardFields.frx":02E0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQuoteBoardFields.frx":030E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQuoteBoardFields.frx":032E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1095
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
         Caption         =   "frmQuoteBoardFields.frx":034A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQuoteBoardFields.frx":0370
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQuoteBoardFields.frx":0390
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDefaults 
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   2400
         Width           =   1095
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
         Caption         =   "frmQuoteBoardFields.frx":03AC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQuoteBoardFields.frx":03DE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQuoteBoardFields.frx":03FE
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgFields 
      Height          =   3555
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   3075
      _cx             =   5424
      _cy             =   6271
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
   Begin HexUniControls.ctlUniLabelXP lblQuoteBoardFields 
      Height          =   195
      Left            =   180
      Top             =   60
      Width           =   2835
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
      Caption         =   "frmQuoteBoardFields.frx":041A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmQuoteBoardFields.frx":0480
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQuoteBoardFields.frx":04A0
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmQuoteBoardFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmQuoteBoardFields.frm
'' Description: Form to allow the user to change the order of their quote
''              board fields, add new ones, and edit or delete existing ones
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/03/2001   DAJ         Created
'' 05/08/2012   DAJ         Added Turnkey mode
'' 10/23/2012   DAJ         Rename Turnkey to HedgeLinc
'' 11/15/2013   DAJ         Changed how to get the Turnkey product name
'' 03/07/2014   DAJ         Moved Cattle stuff into NavCattle.DLL
'' 03/19/2014   DAJ         Renamed Turnkey to Cattle; For Cattle mode, sort
''                          the list and allow for click anywhere on row
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum eQbfMode
    eQbfMode_QBFld = 0
    eQbfMode_CotReport
    eQbfMode_QBCat
    eQbfMode_OptionChain
    eQbfMode_Cattle
    eQbfMode_SaiReport
    eQbfMode_SaiElite
End Enum

Private Enum eGDCols
    eGDCol_Active = 0
    eGDCol_Name
    eGDCol_ID
    eGDCol_Show
    eGDCol_GridStyle
    eGDCol_Detached
    eGDCol_NumFields
End Enum

Private Enum eNewTabMode
    eNewTab_NameOnly = 0        'prompt for name (for adding brand new tab)
    eNewTab_NameIdx             'prompt for name & save index (for copying tab)
End Enum

Private Type mPrivate
    Mode As eQbfMode
    bOK As Boolean
    bMsgShow As Boolean
    astrDefaults As cGdArray
    astrTabsRemoved As cGdArray
    iLastButton As Integer              ' Last mouse button pressed
End Type
Private m As mPrivate

Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAdd_Click
'' Description: If the user clicks on the add button, allow them to create a
''              new quote board field either from scratch or from an existing
''              filter criteria
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAdd_Click()
On Error GoTo ErrSection:

    Dim strText As String
    Dim strID As String                 ' Return from the QBFNew form
    Dim Criteria As New cCriteria       ' Criteria to add
    Dim astrToAdd As New cGdArray       ' Return from the QBFNew form
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Index into a for loop
    Dim strFields As String             ' String to pass to the QBFNew form
    Dim bGrid As Boolean                ' Grid style
    Dim bBox As Boolean                 ' box style

    If m.Mode = eQbfMode_QBFld Then
        'Fields
        'If Not HasGold(True, "Creating custom Quote Board Fields") Then
        '    Exit Sub
        'End If

        strFields = ""
        For lIndex = 1 To fgFields.Rows - 1
            If fgFields.TextMatrix(lIndex, GDCol(eGDCol_Show)) = "True" Then
                strFields = strFields & CStr(CLng(fgFields.RowHidden(lIndex))) & ";" & _
                        fgFields.TextMatrix(lIndex, GDCol(eGDCol_Name)) & ";" & _
                        fgFields.TextMatrix(lIndex, GDCol(eGDCol_ID)) & "|"
            End If
        Next lIndex

        Set astrToAdd = frmQBFNew.ShowMe(strFields)
        If Not astrToAdd Is Nothing Then
            For lIndex = 0 To astrToAdd.Size - 1
                If Left(astrToAdd(lIndex), 1) = "*" Then
                    astrToAdd(lIndex) = Mid(astrToAdd(lIndex), 2)
                    For lRow = 1 To fgFields.Rows - 1
                        If fgFields.TextMatrix(lRow, GDCol(eGDCol_Name)) = astrToAdd(lIndex) Then
                            CheckedCell(fgFields, lRow, GDCol(eGDCol_Active)) = True
                            fgFields.RowHidden(lRow) = False
                            Exit For
                        End If
                    Next lRow
                Else
                    If Criteria.FromFile(App.Path & "\Custom", astrToAdd(lIndex)) Then
                        fgFields.AddItem "1" & vbTab & Criteria.Name & vbTab & Criteria.ID
                        fgFields.Row = fgFields.Rows - 1
                    End If
                End If
                
                If lIndex = 0 And FormIsLoaded("frmAlerts") Then
                    frmAlerts.NewField = astrToAdd(lIndex)
                End If
            Next lIndex
        End If
    ElseIf m.Mode = eQbfMode_SaiReport Then
        ' add symbol(s)
        Set astrToAdd = frmSymbolSelector.ShowMe("")
        If astrToAdd.Size > 0 Then
            For lIndex = 0 To astrToAdd.Size - 1
                strID = Trim(GetSymbol(astrToAdd(lIndex)))
                If Len(strID) > 0 Then
                    strText = g.SymbolPool.Desc(g.SymbolPool.PoolRecForSymbol(strID))
                    fgFields.AddItem "1" & vbTab & strID & vbTab & " " & strText & vbTab & "1"
                End If
            Next
            'fgFields.Select fgFields.Rows - 1, 1
            fgFields.ShowCell fgFields.Rows - 1, 1
        End If
    ElseIf m.Mode = eQbfMode_SaiElite Then
        ' add symbol(s)
        Set astrToAdd = frmSymbolSelector.ShowMe("")
        If astrToAdd.Size > 0 Then
            For lIndex = 0 To astrToAdd.Size - 1
                strID = Trim(GetSymbol(astrToAdd(lIndex)))
                If Len(strID) > 0 Then
                    strText = g.SymbolPool.Desc(g.SymbolPool.PoolRecForSymbol(strID))
                    For lRow = fgFields.FixedRows To fgFields.Rows - 1
                        If strID < fgFields.TextMatrix(lRow, 1) Then
                            Exit For
                        End If
                    Next
                    fgFields.AddItem "1" & vbTab & strID & vbTab & " " & strText & vbTab & "1", lRow
                End If
            Next
            'fgFields.Select fgFields.Rows - 1, 1
            fgFields.ShowCell lRow, 1 'fgFields.Rows - 1, 1
        End If
    Else
        'Categories
        If Not HasGold(True, "Creating new Quote Board tabs") Then
            Exit Sub
        End If
        
        ProcessNewTab eNewTab_NameOnly
    End If
    
    EnableButtons
    
ErrExit:
    Set Criteria = Nothing
    Set astrToAdd = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.cmdAdd.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user clicks on the cancel, set the OK to True and hide
''              the form to allow the ShowMe to continue
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
    RaiseError "frmQuoteBoardFields.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdDefaults_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    
    
    If cmdDefaults.Caption = "Co&py" Then
        ProcessNewTab eNewTab_NameIdx
    Else
        ' If the defaults array is nothing for some reason, then bail...
        If m.astrDefaults Is Nothing Then Exit Sub
        
        With fgFields
            .Redraw = flexRDNone
            
            .Rows = .FixedRows
            For lIndex = 0 To m.astrDefaults.Size - 1
                .AddItem m.astrDefaults(lIndex)
                .RowHidden(lIndex + 1) = Not CBool(.TextMatrix(lIndex + 1, GDCol(eGDCol_Show)))
                If m.Mode = eQbfMode_QBFld Then
                    .RowHidden(lIndex + 1) = Not CheckedCell(fgFields, lIndex + 1, GDCol(eGDCol_Active))
                End If
            Next lIndex
            
            .Redraw = flexRDBuffered
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.cmdDefaults.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDown_Click
'' Description: If the user clicks on the "Move Down" button, move the selected
''              row down one row
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDown_Click()
On Error GoTo ErrSection:

    Dim lRow As Long

    With fgFields
        If .RowSel > .FixedRows - 1 And .RowSel < .Rows - 1 Then
            lRow = .RowSel + 1
            If lRow + 1 < .Rows Then
                Do While (.RowHidden(lRow + 1) Or .RowHidden(lRow)) And lRow < .Rows - 1
                    lRow = lRow + 1
                    If lRow = .Rows - 1 Then Exit Do
                Loop
            End If
            
            .RowPosition(.RowSel) = lRow
            .Row = lRow
            .RowSel = lRow
        End If
    End With
    
    EnableButtons

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.cmdDown.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: If the user clicks on the edit button, bring up the criteria
''              editor for the quote board field that they have selected
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    Dim strText As String
    Dim frm As frmCriteria
    Dim strID As String
    Dim Criteria As New cCriteria
    
    With fgFields
        If m.Mode = eQbfMode_QBFld Then
            Set frm = New frmCriteria
            strID = frm.ShowMe(AddSlash(App.Path) & "Custom\", .TextMatrix(.Row, GDCol(eGDCol_ID)), True, eCriteria_FilterCriteria)
            If strID <> "" Then
                If Criteria.FromFile(AddSlash(App.Path) & "Custom\", strID) Then
                    .TextMatrix(.Row, GDCol(eGDCol_ID)) = strID
                    .TextMatrix(.Row, GDCol(eGDCol_Name)) = Criteria.Name
                End If
            End If
        Else
            strText = .TextMatrix(.Row, GDCol(eGDCol_Name))
            strText = Trim(InfBox("New name for tab:", "?", , "Rename Tab", , , , , , "s", strText))
            If Len(strText) > 0 Then
                If UCase(Trim(strText)) = "(FILTER)" Then
                    strText = strText & " is a reserved name." & vbCrLf & "Please supply a different name"     '4238
                    InfBox strText, "E"
                Else
                    .TextMatrix(.Row, GDCol(eGDCol_Name)) = strText
                End If
            End If
        End If
    End With
    
ErrExit:
    Set frm = Nothing
    Set Criteria = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.cmdEdit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user clicks on OK, set the OK flag to True and hide the
''              form to allow the ShowMe to continue
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If m.Mode = eQbfMode_QBCat Then
        m.bOK = OkayToDetach
        If Not m.bOK Then InfBox "You cannot detach all quote tabs.", "I"
    Else
        m.bOK = True
    End If
    
    If m.bOK Then Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemove_Click
'' Description: If the user clicks on remove, remove the item from the quote
''              board and delete the QBF file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemove_Click()
On Error GoTo ErrSection:

    Dim strReturn As String             ' Whether or not the user wants to delete
    
    With fgFields
        If m.Mode = eQbfMode_QBCat Then
            If Not HasGold(True, "Removing Quote Board Tabs") Then Exit Sub
            If .Rows = .FixedRows + 1 Then
                Err.Raise vbObjectError + 1000, , "There must be at least one tab on the quote board"
            End If
        End If
        If m.Mode = eQbfMode_QBFld Then
            strReturn = AskBox("h=Confirmation ; i=? ; b=+Remove|-Cancel ; " & _
                "Are you sure that you want to remove " & .TextMatrix(.Row, GDCol(eGDCol_Name)) & "?")
        End If
        If strReturn <> "C" Then
            If m.Mode = eQbfMode_QBFld Then
                If .TextMatrix(.Row, GDCol(eGDCol_ID)) = "" Then
                    CheckedCell(fgFields, .Row, GDCol(eGDCol_Active)) = False
                    .RowHidden(.Row) = True
                Else
                    .RemoveItem .Row
                End If
            Else
                'if removing a QB tab then make sure associated tab alerts are also removed
                If m.Mode = eQbfMode_QBCat Then
                    m.astrTabsRemoved.Add .TextMatrix(.Row, GDCol(eGDCol_Name))
                End If
                .RemoveItem .Row
            End If
        End If
    End With
    
    EnableButtons
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.cmdRemove.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdUp_Click
'' Description: If the user clicks on the "Move Up" button, move the selected
''              row up one row
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdUp_Click()
On Error GoTo ErrSection:

    Dim lRow As Long

    With fgFields
        If .RowSel > .FixedRows Then
            lRow = .RowSel - 1
            If lRow - 1 > .FixedRows Then
                Do While (.RowHidden(lRow - 1) Or .RowHidden(lRow)) And lRow > .FixedRows
                    lRow = lRow - 1
                    If lRow = .FixedRows Then Exit Do
                Loop
            End If
            
            .RowPosition(.RowSel) = lRow
            .Row = lRow
            .RowSel = lRow
        End If
    End With
    
    EnableButtons

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.cmdUp.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFields_AfterMoveRow
'' Description: After the user has moved a row, make sure that it is selected
'' Inputs:      Row moved, New Position of row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFields_AfterMoveRow(ByVal Row As Long, Position As Long)
On Error GoTo ErrSection:

    With fgFields
        .Row = Position
        .RowSel = Position
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.fgFields.AfterMoveRow", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFields_BeforeEdit
'' Description: Only allow the user to edit the "Active" field
'' Inputs:      Row and Column being edited, Whether or not to cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFields_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim strText$

    If m.Mode = eQbfMode_QBCat Then
        With fgFields
            If .Row >= .FixedRows And .Row < .Rows Then
                strText = .TextMatrix(.Row, GDCol(eGDCol_ID))
                If Len(strText) > 0 Then
                    If Left(strText, 1) = "-" Then
                        If cmdDefaults.Enabled Then cmdDefaults.Enabled = False
                    ElseIf Not cmdDefaults.Enabled Then
                        cmdDefaults.Enabled = True
                    End If
                ElseIf cmdDefaults.Enabled Then
                    cmdDefaults.Enabled = False
                End If
            End If
        End With
    End If
    
    If Col <> GDCol(eGDCol_Active) And Col <> GDCol(eGDCol_Detached) Then
        Cancel = True
    ElseIf Col = GDCol(eGDCol_Detached) Then
'JM: 08-28-2009 - 5282 (Per Pete, allow Better Trades to detached QB)
        If ExtremeCharts = 0 Then
            If Not HasLevel(eTN3_Standard, m.bMsgShow, "Detaching quote tabs ") Then
                Cancel = True
            End If
        End If
        m.bMsgShow = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.fgFields.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFields_BeforeMouseDown
'' Description: Mark the row as a drag row to allow the user to move it where
''              they want it
'' Inputs:      Mouse button pressed, Shift/Ctrl/Alt status, Location of click,
''              Whether or not to cancel the operation
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFields_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position of new row
    Dim lRow As Long                    ' Row being moved
    
    m.iLastButton = Button
    With fgFields
        lRow = .MouseRow
        If (lRow <> -1) And (fraUpDown.Visible = True) Then
            .Row = lRow
            .RowSel = lRow
            
            .Refresh
            lPos = .DragRow(lRow)
            If lPos <> lRow Then
                Cancel = True
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.fgFields.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFields_Click
'' Description: Make sure to select the row that the user clicks on
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFields_Click()
On Error GoTo ErrSection:

    With fgFields
        If .MouseRow <> -1 Then
            If m.iLastButton = vbLeftButton Then
                .Row = .MouseRow
                .RowSel = .Row
            
                ' DAJ 03/18/2014: If the user clicks the left mouse button on a valid row
                ' in Turnkey mode, toggle the active column regardless of where they clicked...
                If m.Mode = eQbfMode_Cattle Then
                    If ValidGridRow(fgFields, .Row) Then
                        CheckedCell(fgFields, .Row, GDCol(eGDCol_Active)) = Not CheckedCell(fgFields, .Row, GDCol(eGDCol_Active))
                    End If
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.fgFields.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgFields_AfterRowColChange
'' Description: Enable/Disable the buttons according to what row the user is
''              currently on
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgFields_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    EnableButtons

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.fgFields.AfterRowColChange", eGDRaiseError_Show
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
    RaiseError "frmQuoteBoardFields.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, center it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16(ToolbarIcon("kSelect"))
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: When the user closes the form, cancel the unload and set the
''              OK to False so that the ShowMe can continue
'' Inputs:      Whether or not to cancel the unload, Unload Mode
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
    RaiseError "frmQuoteBoardFields.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: If the user resizes the form, move and resize the controls on
''              the form appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lTop As Long                    ' Top of a control
    Dim lHeight As Long                 ' Height of a control

    If LimitFormSize(Me, fraButtons.Width * 4, fraButtons.Height + fraUpDown.Height) = False Then
        With fgFields
            If lblQuoteBoardFields.Visible Then
                lTop = .Top
            Else
                lTop = fraButtons.Top
            End If
            lHeight = ScaleHeight - lTop - 120
            If fraUpDown.Visible Then
                lHeight = lHeight - fraUpDown.Height
            End If
            
            .Move .Left, lTop, ScaleWidth - fraButtons.Width - (.Left * 3), lHeight
            
            If .ColHidden(eGDCol_Detached) = False Then
                .ColWidth(eGDCol_Name) = .ClientWidth - 900
                .ColWidth(eGDCol_Detached) = 900
            End If
        End With
        
        With fraButtons
            .Move fgFields.Width + fgFields.Left * 2
        End With
        
        With fraUpDown
            .Move (fgFields.Width / 2) - (.Width / 2) + fgFields.Left, ScaleHeight - .Height
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initializes the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long

    With fgFields
        lRedraw = .Redraw
        .Redraw = flexRDNone
        SetupGrid fgFields, eGridMode_Grid
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .GridLines = flexGridNone
        .GridLinesFixed = flexGridInset
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumFields)
        .FixedCols = 0
        .RowHidden(0) = True
        
        .TextMatrix(0, GDCol(eGDCol_Active)) = "Show"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_ID)) = "ID"
        .TextMatrix(0, GDCol(eGDCol_Show)) = "Show"
        .TextMatrix(0, GDCol(eGDCol_Detached)) = "Detached"
        
        .ColDataType(GDCol(eGDCol_Active)) = flexDTBoolean
        .ColDataType(GDCol(eGDCol_Show)) = flexDTBoolean
        .ColDataType(GDCol(eGDCol_Detached)) = flexDTBoolean
        
        .ColHidden(GDCol(eGDCol_ID)) = True
        .ColHidden(GDCol(eGDCol_Show)) = True
        .ColHidden(GDCol(eGDCol_GridStyle)) = True
        .ColHidden(GDCol(eGDCol_Detached)) = True
        
        If m.Mode = eQbfMode_SaiReport Or m.Mode = eQbfMode_SaiElite Then
            .ColHidden(GDCol(eGDCol_ID)) = False
            .ColHidden(GDCol(eGDCol_Active)) = True 'False
        ElseIf m.Mode = eQbfMode_QBCat Or m.Mode = eQbfMode_QBFld Then
            .ColHidden(GDCol(eGDCol_Active)) = True
            If m.Mode = eQbfMode_QBCat Then .ColHidden(GDCol(eGDCol_Detached)) = False
        Else
            .ColHidden(GDCol(eGDCol_Active)) = False
        End If
        
        If m.Mode = eQbfMode_QBFld Then
            .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbWindowBackground
            .GridLines = flexGridNone
            .GridLinesFixed = flexGridNone
        End If
        
        .AutoSize 0
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.InitGrid", eGDRaiseError_Raise
   
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Fill in the grid with information passed in, show the form,
''              and pass back information from the grid if the user clicks on OK
'' Inputs:      Array of tab delimited strings with the fields of the grid,
''              Mode to show the form in
'' Returns:     True if OK, False if Cancel
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(astrFields As cGdArray, _
        Mode As eQbfMode, _
        Optional astrDefaults As cGdArray = Nothing, _
        Optional ByVal bCopy As Boolean = False, _
        Optional ByVal strStart As String = "") As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bOkNew As Boolean
    Dim strName As String
    Dim strIdx As String

    m.Mode = Mode
    m.bMsgShow = True
    Set m.astrDefaults = astrDefaults
    Set m.astrTabsRemoved = New cGdArray
           
    InitGrid
    
    ' Although this should theoretically never happen, at times customer end up with a blank field
    ' in this array which causes a problem in the next loop, so we will walk through and remove
    ' blanks here.  2/8/2007 DAJ...
    For lIndex = astrFields.Size - 1 To 0 Step -1
        If Len(astrFields(lIndex)) = 0 Then astrFields.Remove lIndex
    Next lIndex
        
    For lIndex = 0 To astrFields.Size - 1
        With fgFields
            .AddItem astrFields(lIndex)
            .RowHidden(lIndex + 1) = Not CBool(fgFields.TextMatrix(lIndex + 1, GDCol(eGDCol_Show)))
            If .RowHidden(lIndex + 1) = False And m.Mode = eQbfMode_QBFld Then
                .RowHidden(lIndex + 1) = Not CheckedCell(fgFields, lIndex + 1, GDCol(eGDCol_Active))
            End If
            If m.Mode = eQbfMode_QBCat Then
                'astrFields format from QB
                '0:active \t 1:name \t 2:QB table index \t 3:show \t 4:style flag \t 5:detached
                If .TextMatrix(.Rows - 1, GDCol(eGDCol_Detached)) = "0" Then
                    .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Detached)) = flexUnchecked
                Else
                    .Cell(flexcpChecked, .Rows - 1, GDCol(eGDCol_Detached)) = flexChecked
                End If
            End If
        End With
    Next lIndex
    
    With fgFields
        If m.Mode = eQbfMode_Cattle Then
            .Col = GDCol(eGDCol_Name)
            .Sort = flexSortStringAscending
        End If
        
        If Len(strStart) > 0 Then
            For lIndex = .FixedRows To .Rows - 1
                If .TextMatrix(lIndex, GDCol(eGDCol_Name)) = strStart Then
                    .Row = lIndex
                    .RowSel = lIndex
                    Exit For
                End If
            Next lIndex
        Else
            If .Rows > .FixedRows Then
                .Row = .FixedRows
                .RowSel = .Row
            End If
        End If
    End With
    
    Select Case Mode
        Case eQbfMode_QBFld
            Caption = "Quote Board Fields"
            cmdAdd.Visible = True
            cmdRemove.Visible = True
            cmdEdit.Visible = True
            cmdEdit.Caption = "&Edit"
            cmdDefaults.Visible = True
            cmdDefaults.Caption = "De&faults"
            cmdUp.Visible = True
            cmdDown.Visible = True
            lblQuoteBoardFields.Visible = True
        
        Case eQbfMode_QBCat
            Caption = "Quote Board Tabs"
            cmdAdd.Visible = True
            cmdRemove.Visible = True
            cmdEdit.Visible = True
            cmdEdit.Caption = "Re&name"
            cmdDefaults.Visible = True
            cmdDefaults.Caption = "Co&py"
            fraUpDown.Visible = True
            lblQuoteBoardFields.Visible = False      'hide label that says "fields for current quote tab"
            With fgFields
                'some users name there quote-tab with leading numeric (eg Bill Lopinto names his 60-minute tab as 60)
                .Cell(flexcpAlignment, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
                .RowHidden(.FixedRows - 1) = False
            End With
            
            If bCopy Then
                If Not ProcessNewTab(eNewTab_NameIdx) Then
                    m.bOK = False
                    ShowMe = False
                    GoTo ErrExit
                End If
            End If
        
        Case eQbfMode_CotReport
            Caption = "COT Report Fields"
            cmdAdd.Visible = False
            cmdRemove.Visible = False
            cmdEdit.Visible = False
            cmdDefaults.Visible = False
            fraUpDown.Visible = False
            lblQuoteBoardFields.Visible = False
        
        Case eQbfMode_OptionChain
            Caption = "Option Chain Fields"
            cmdAdd.Visible = False
            cmdRemove.Visible = False
            cmdEdit.Visible = False
            cmdDefaults.Top = cmdAdd.Top
            cmdDefaults.Visible = True
            cmdDefaults.Caption = "De&faults"
            fraUpDown.Visible = True
            lblQuoteBoardFields.Visible = False
    
        Case eQbfMode_Cattle
            Caption = g.CattleBridge.ProductName & " Fields"
            cmdAdd.Visible = False
            cmdRemove.Visible = False
            cmdEdit.Visible = False
            cmdDefaults.Visible = False
            fraUpDown.Visible = False
            lblQuoteBoardFields.Visible = False
            fgFields.Editable = flexEDNone
            
        Case eQbfMode_SaiReport
            Caption = "Symbols for the SAI Report"
            cmdAdd.Visible = True
            cmdRemove.Visible = True
            cmdEdit.Visible = False
            cmdDefaults.Visible = True
            fraUpDown.Visible = True
            lblQuoteBoardFields.Visible = False
        
        Case eQbfMode_SaiElite
            Caption = "Symbols for the SAI Elite Report"
            cmdAdd.Visible = True
            cmdRemove.Visible = True
            cmdEdit.Visible = False
            cmdDefaults.Visible = True
            fraUpDown.Visible = False 'True
            lblQuoteBoardFields.Visible = False
    End Select
    
    EnableButtons
    ShowForm Me, eForm_ActModal
    
    If m.bOK = True Then
        RemoveTabAlerts
        
        astrFields.Clear
        
        With fgFields
            For lIndex = .FixedRows To .Rows - 1
                If Mode = eQbfMode_QBCat Then
                    strName = .TextMatrix(lIndex, GDCol(eGDCol_Name))
                    strIdx = .TextMatrix(lIndex, GDCol(eGDCol_ID))
                    
                    If Len(strIdx) = 0 Then
                        strName = Parse(strName, "(add pending)", 1)
                        strIdx = "-999999"          '5797
                    ElseIf Left(strIdx, 1) = "-" Then
                        strName = Parse(strName, "(copy pending)", 1)
                    End If
                    
                    'active \t name \t QB table index \t show \t style \t detached
                    astrFields.Add CStr(.Cell(flexcpChecked, lIndex, GDCol(eGDCol_Active))) & _
                        vbTab & strName & vbTab & strIdx & vbTab & _
                        CStr(.TextMatrix(lIndex, GDCol(eGDCol_Show))) & _
                        vbTab & .TextMatrix(lIndex, GDCol(eGDCol_GridStyle)) & vbTab & .Cell(flexcpChecked, lIndex, GDCol(eGDCol_Detached))
                Else
                    astrFields.Add CStr(.Cell(flexcpChecked, lIndex, GDCol(eGDCol_Active))) & _
                        vbTab & .TextMatrix(lIndex, GDCol(eGDCol_Name)) & vbTab & _
                        .TextMatrix(lIndex, GDCol(eGDCol_ID)) & vbTab & CStr(.TextMatrix(lIndex, GDCol(eGDCol_Show)))
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
    RaiseError "frmQuoteBoardFields.ShowMe", eGDRaiseError_Raise
    
End Function

Private Sub EnableButtons()
On Error GoTo ErrSection:

    With fgFields
        If .Row < .FixedRows Or .Row >= .Rows Then
            cmdUp.Enabled = False
            cmdDown.Enabled = False
            cmdEdit.Enabled = False
            cmdRemove.Enabled = False
        Else
            cmdUp.Enabled = (.Row > .FixedRows)
            cmdDown.Enabled = (.Row > .FixedRows - 1 And LastVisibleRow(.Row) = False)
            If m.Mode = eQbfMode_QBFld Then
                cmdEdit.Enabled = (.TextMatrix(.Row, GDCol(eGDCol_ID)) <> "")
                cmdRemove.Enabled = True '(.TextMatrix(.Row, GDCol(eGDCol_ID)) <> "")
            Else
                cmdEdit.Enabled = True
                cmdRemove.Enabled = True
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQuoteBoardFields.EnableButtons", eGDRaiseError_Raise
    
End Sub

Private Function LastVisibleRow(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long
    
    LastVisibleRow = True
    If lRow < fgFields.Rows - 1 Then
        For lIndex = lRow + 1 To fgFields.Rows - 1
            If Not fgFields.RowHidden(lIndex) Then
                LastVisibleRow = False
                Exit For
            End If
        Next lIndex
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmQuoteBoardFields.LastVisibleRow", eGDRaiseError_Raise
    
End Function

Private Function OkayToDetach() As Boolean
On Error GoTo ErrSection:

    Dim i&
    Dim bOkay As Boolean
    
    bOkay = frmQuotes.HasFilterTab
    
    If Not bOkay Then
        With fgFields
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, eGDCol_Detached) = flexUnchecked Then
                    bOkay = True
                    Exit For
                End If
            Next
        End With
    End If
    
    OkayToDetach = bOkay
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmQuoteBoardFields.OkayToDetach"

End Function

Private Sub RemoveTabAlerts()
On Error GoTo ErrSection:

    Dim i, j, strTab$
    
    Dim Alert As cAlert
    Dim bUpdateForm As Boolean


    If m.astrTabsRemoved Is Nothing Then Exit Sub
    
    For i = 0 To m.astrTabsRemoved.Size - 1
        strTab = m.astrTabsRemoved(i)
        For j = g.Alerts.Count To 1 Step -1
            Set Alert = g.Alerts(j)
            If Not Alert Is Nothing Then
                If Alert.TabName = strTab Then
                    g.Alerts.Remove Alert.AlertKey
                    bUpdateForm = True
                End If
            End If
        Next
    Next
    
    If bUpdateForm And FormIsLoaded("frmAlertsSetup") Then
        frmAlertsSetup.LoadGrid
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmQuoteBoardFields.RemoveTabALerts"

End Sub

Private Function ProcessNewTab(ByVal eMode As eNewTabMode) As Boolean
On Error GoTo ErrSection:

    Dim bGrid As Boolean
    Dim bBox As Boolean
    Dim eQbStyle As eGDQuoteStyle
    
    Dim i&, strName$, strTemp$
    Dim aNew As New cGdArray
    
    If fgFields.Row < fgFields.FixedRows Or fgFields.Row >= fgFields.Rows Then
        i = frmQuotes.CurrentTabNum(True)
        strName = UCase(frmQuotes.CurrentTabName(i))
        If strName = "(FILTER)" Then
            strTemp = i
        Else
            Exit Function
        End If
    Else
        strTemp = fgFields.TextMatrix(fgFields.Row, GDCol(eGDCol_ID))
        i = ValOfText(strTemp)
    End If
    
    If eMode = eNewTab_NameIdx Then
        If Len(strTemp) = 0 Or i < 0 Then Exit Function
        strTemp = "-" & strTemp
    ElseIf eMode = eNewTab_NameOnly Then
        strTemp = ""
        If i <= 0 Then i = frmQuotes.CurrentTabNum(True)
        If i < 0 Then Exit Function
    End If
    
    eQbStyle = frmQuotes.TabStr(eGDTabSettings_Style, i)
    
    If eQbStyle = eGDQuoteStyle_Grid Then
        bGrid = True
    ElseIf eQbStyle <> eGDQuoteStyle_Forex Then
        bBox = True
    End If
    
    If frmPassword.ShowNewTab(strName, bGrid, bBox) Then
        If eMode = eNewTab_NameOnly Then
            strName = strName & " (add pending)"
        Else
            strName = strName & " (copy pending)"
        End If
        
        If bGrid Then
            eQbStyle = eGDQuoteStyle_Grid
        ElseIf bBox Then
            eQbStyle = eGDQuoteStyle_OHLC
        Else
            eQbStyle = eGDQuoteStyle_Forex
        End If
        
        'active \t name \t tabindex \t show \t style
        fgFields.AddItem "True" & vbTab & strName & vbTab & strTemp & vbTab & "True" & vbTab & Str(eQbStyle)
        fgFields.Row = fgFields.Rows - 1
        fgFields.RowSel = fgFields.Rows - 1
        
        If cmdDefaults.Enabled Then cmdDefaults.Enabled = False
        
        ProcessNewTab = True
    End If
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmQuoteBoardFields.ProcessNewTab"

End Function

