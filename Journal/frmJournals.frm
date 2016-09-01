VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmJournals 
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMenu 
      Left            =   120
      Top             =   3780
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   3600
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
      Caption         =   "frmJournals.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmJournals.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmJournals.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdPrint 
         Height          =   495
         Left            =   1320
         TabIndex        =   11
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
         Caption         =   "frmJournals.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmJournals.frx":0094
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmJournals.frx":00B4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdClose 
         Height          =   495
         Left            =   0
         TabIndex        =   10
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
         Caption         =   "frmJournals.frx":00D0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmJournals.frx":00FC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmJournals.frx":011C
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraFilter 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6915
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
      Caption         =   "frmJournals.frx":0138
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmJournals.frx":0182
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmJournals.frx":01A2
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkToDate 
         Height          =   220
         Left            =   3540
         TabIndex        =   7
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmJournals.frx":01BE
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmJournals.frx":01EE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmJournals.frx":020E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkFromDate 
         Height          =   220
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmJournals.frx":022A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmJournals.frx":0260
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmJournals.frx":0280
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkSymbol 
         Height          =   220
         Left            =   3540
         TabIndex        =   3
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmJournals.frx":029C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmJournals.frx":02CC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmJournals.frx":02EC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkAccount 
         Height          =   220
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmJournals.frx":0308
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmJournals.frx":033A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmJournals.frx":035A
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboSymbols 
         Height          =   315
         Left            =   4680
         TabIndex        =   4
         Top             =   240
         Width           =   2115
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   -2147483630
         ButtonStyle     =   -1
         SelectorStyle   =   -1
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
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
         Tip             =   "frmJournals.frx":0376
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmJournals.frx":0396
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboAccounts 
         Height          =   315
         Left            =   1260
         TabIndex        =   2
         Top             =   240
         Width           =   2115
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   -2147483630
         ButtonStyle     =   -1
         SelectorStyle   =   -1
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
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
         Tip             =   "frmJournals.frx":03B2
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmJournals.frx":03D2
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectDate gdFromDate 
         Height          =   315
         Left            =   1260
         TabIndex        =   6
         Top             =   660
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
      End
      Begin gdOCX.gdSelectDate gdToDate 
         Height          =   315
         Left            =   4680
         TabIndex        =   8
         Top             =   660
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgJournal 
      Height          =   2115
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   6915
      _cx             =   12197
      _cy             =   3731
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
      Caption         =   "Pop Up"
      Begin VB.Menu mnuEditJournal 
         Caption         =   "Edit Journal Entry"
      End
      Begin VB.Menu mnuDeleteJournal 
         Caption         =   "Delete Journal Entry"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmJournals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmJournals.frm
'' Description: Allows the user to view all journals with filters
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/20/2009   DAJ         Added a column for the emotion number
'' 04/22/2009   DAJ         Added Horz scroll, show blank if EmotionNumber = -1
'' 04/23/2009   DAJ         Changed from database to new journal objects
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 09/08/2014   DAJ         Use NavCore Image List; Use newer place/save form
'' 10/24/2014   DAJ         Core Application functions for DLL's
'' 05/18/2015   DAJ         Pass frmPrintPreview.vp to DoPrintHeader
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_JournalID = 0
    eGDCol_Date
    eGDCol_OrderID
    eGdCol_BuySell
    eGDCol_Symbol
    eGDCol_Account
    eGDCol_Action
    eGDCol_EmotionNumber
    eGDCol_Feelings
    eGDCol_Reasons
    eGDCol_Thoughts
    eGDCol_Note
    eGDCol_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on the OK button?
    Journals As cJournals               ' Collection of order journals
End Type
Private m As mPrivate

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup, load, and show the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe()
On Error GoTo ErrSection:
    
    Load
    
    If g.bAppIsIde Then
        mGenesis.ShowForm Me, eForm_Modal
    Else
        g.TnCore.ShowForm Me, eForm_Modal
    End If

ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    g.TnCore.RaiseError "frmJournals.ShowMe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow the user to print the journal
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection:

    PrintMe = frmPrintPreview.ShowMe("CNV Journals", frmJournals, , 0.75, 0.75, 0.75, 0.75, True)

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.PrintMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the print preview control
'' Inputs:      Args to pass to print preview
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    With frmPrintPreview.vp
        .StartDoc
        g.TnCore.DoPrintHeader , frmPrintPreview.vp
        
        .FontName = "Times New Roman"
        .FontSize = 14
        .FontBold = True
        
        .Text = "Order Journals" & vbCrLf
        
        .FontSize = 12
        .FontBold = False
        
        .Text = "Filters:" & vbLf
        If chkAccount.Value = vbChecked Then
            .Text = "Account: " & cboAccounts.Text
        Else
            .Text = "Account: ALL"
        End If
        If chkSymbol.Value = vbChecked Then
            .Text = "; Symbol: " & cboSymbols.Text
        Else
            .Text = "; Symbol: ALL"
        End If
        If chkFromDate.Value = vbChecked Then
            .Text = "; From Date: " & DateFormat(gdFromDate.Value, MM_DD_YY)
        Else
            .Text = "; From Date: NONE"
        End If
        If chkToDate.Value = vbChecked Then
            .Text = "; To Date: " & DateFormat(gdToDate.Value, MM_DD_YY)
        Else
            .Text = "; To Date: NONE"
        End If
        
        .Text = vbCrLf
        
        .Paragraph = ""
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable fgJournal
        Else
            fgJournal.ColWidth(GDCol(eGDCol_Note)) = 3000
            .RenderControl = fgJournal.hWnd
            AutoSizeGrid
        End If
        .Paragraph = ""

        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.GenerateReport"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboAccounts_Click
'' Description: If the user changes the value, refilter the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAccounts_Click()
On Error GoTo ErrSection:

    If (Visible = True) And (chkAccount.Value = vbChecked) Then
        FilterGrid
        SaveFilters
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.cboAccounts_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboSymbols_Click
'' Description: If the user changes the value, refilter the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboSymbols_Click()
On Error GoTo ErrSection:

    If (Visible = True) And (chkSymbol.Value = vbChecked) Then
        FilterGrid
        SaveFilters
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.cboSymbols_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkAccount_Click
'' Description: If the user changes the value, refilter the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAccount_Click()
On Error GoTo ErrSection:

    If Visible Then
        FilterGrid
        SaveFilters
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.chkAccount_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkFromDate_Click
'' Description: If the user changes the value, refilter the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkFromDate_Click()
On Error GoTo ErrSection:

    If Visible Then
        FilterGrid
        SaveFilters
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.chkFromDate_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkSymbol_Click
'' Description: If the user changes the value, refilter the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkSymbol_Click()
On Error GoTo ErrSection:

    If Visible Then
        FilterGrid
        SaveFilters
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.chkSymbol_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkToDate_Click
'' Description: If the user changes the value, refilter the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkToDate_Click()
On Error GoTo ErrSection:

    If Visible Then
        FilterGrid
        SaveFilters
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.chkToDate_Click"
    
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

    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.cmdClose_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPrint_Click
'' Description: Allow the user to print the journals
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
    g.TnCore.RaiseError "frmJournals.cmdPrint_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgJournal_AfterSort
'' Description: After sorting the grid, recolor rows and save last sort
'' Inputs:      Column of Sort, Order of Sort
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgJournal_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SetBackColors fgJournal
    
    SetIniFileProperty "Sorting", Str(Col) & ";" & Str(Order), "Journals", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.fgJournal_AfterSort"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgJournal_BeforeMouseDown
'' Description: If the user right clicks on the grid, bring up the pop up
'' Inputs:      Button, Shift/Ctrl/Alt Status, X, Y, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgJournal_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    fgJournal.Row = fgJournal.MouseRow
    
    If Button = vbRightButton Then
        Enable mnuEditJournal, (fgJournal.Row >= fgJournal.FixedRows) And (fgJournal.Row < fgJournal.Rows)
        Enable mnuDeleteJournal, (fgJournal.Row >= fgJournal.FixedRows) And (fgJournal.Row < fgJournal.Rows)
        
        PopupMenu mnuPopUp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.fgJournal_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgJournal_DblClick
'' Description: Handle user double click in the grid
'' Inputs:      Key Code, Shift/Ctrl/Alt Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgJournal_DblClick()
On Error GoTo ErrSection:

    tmrMenu.Tag = "EditJournal"
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.fgJournal_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgJournal_KeyUp
'' Description: Handle user key presses in the grid
'' Inputs:      Key Code, Shift/Ctrl/Alt Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgJournal_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Select Case KeyCode
        Case vbKeyDelete
            DeleteJournal
            
        Case vbKeyReturn
            tmrMenu.Tag = "EditJournal"
            tmrMenu.Enabled = True
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.fgJournal_KeyUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Do inialization when form is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strFont As String               ' Font to be used for the grid

    Icon = g.CoreBridge.Picture16(g.TnCore.ToolbarIcon("ID_TradeTracker"))
    Caption = "Order Journals"
    
    g.Styler.StyleForm Me
    
    PlaceTheForm Me, g.strIniFile

    strFont = GetIniFileProperty("Journals", "", "Fonts", g.strIniFile)
    If Len(strFont) > 0 Then FontFromString fgJournal.Font, strFont
    
    mnuPopUp.Visible = False
    
    tmrMenu.Enabled = False
    tmrMenu.Interval = 10
    
    Set m.Journals = New cJournals
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: When the user clicks on the X, let ShowMe unload the form
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        Hide
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.Form_QueryUnload"

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

#If 0 Then
    If Not LimitFormSize(Me, (fraFilter.Width + fraSorting.Width) + 360, fraFilter.Height + fraButtons.Height + 360) Then
        With fraFilter
            .Move 120, 120
        End With
        
        With fraButtons
            .Move 120, fraFilter.Height + 240
        End With
        
        With fgJournal
            .Move fraFilter.Width + 240, fraSorting.Height + 240, ScaleWidth - fraFilter.Width - 360, ScaleHeight - fraSorting.Height - 360
        End With
        AutoSizeGrid
    End If
#Else
    If Not LimitFormSize(Me, fraFilter.Width + 240, fraFilter.Height * 3) Then
        With fraFilter
            .Move 120, 120
        End With
        
        With fgJournal
            .Move 120, fraFilter.Height + 240, ScaleWidth - 240, ScaleHeight - fraFilter.Height - fraButtons.Height - 480
        End With
        AutoSizeGrid
        
        With fraButtons
            .Move (ScaleWidth - .Width) / 2, ScaleHeight - .Height - 120
        End With
    End If
#End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up and save settings when the form is unloaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    tmrMenu.Enabled = False
    
    SaveTheFormPlacement Me, g.strIniFile
    SetIniFileProperty "Journals", FontToString(fgJournal.Font), "Fonts", g.strIniFile
    
    SaveFilters
    
    Set m.Journals = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdFromDate_Changed
'' Description: If the user changes the value, refilter the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdFromDate_Changed()
On Error GoTo ErrSection:

    If (Visible = True) And (chkFromDate.Value = vbChecked) Then
        FilterGrid
        SaveFilters
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.gdFromDate_Changed"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdToDate_Changed
'' Description: If the user changes the value, refilter the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdToDate_Changed()
On Error GoTo ErrSection:

    If (Visible = True) And (chkToDate.Value = vbChecked) Then
        FilterGrid
        SaveFilters
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.gdToDate_Changed"

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

    g.TnCore.ChangeGridFont fgJournal, True

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.mnuChangeFont_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuDeleteJournal_Click
'' Description: Allow the user to delete a journal entry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuDeleteJournal_Click()
On Error GoTo ErrSection:

    DeleteJournal

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.mnuDeleteJournal_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditJournal_Click
'' Description: Allow the user to edit a journal entry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEditJournal_Click()
On Error GoTo ErrSection:

    tmrMenu.Tag = "EditJournal"
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.mnuEditJournal_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPrint_Click
'' Description: Allow the user to print the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPrint_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.mnuPrint_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrMenu_Timer
'' Description: Run the stuff from the menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrMenu_Timer()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Current row in the grid
    Dim Journal As cJournal             ' Journal object

    tmrMenu.Enabled = False
    
    Select Case UCase(tmrMenu.Tag)
        Case "EDITJOURNAL"
            If (fgJournal.Row >= fgJournal.FixedRows) And (fgJournal.Row < fgJournal.Rows) Then
                Set Journal = m.Journals.Item(fgJournal.TextMatrix(fgJournal.Row, GDCol(eGDCol_JournalID)))
                
                If frmJournal.ShowMe(fgJournal.RowData(fgJournal.Row), CLng(Val(fgJournal.TextMatrix(fgJournal.Row, GDCol(eGDCol_JournalID)))), Journal) = True Then
                    lRow = fgJournal.Row
                    Load
                    fgJournal.Row = lRow
                    fgJournal.ShowCell lRow, 0&
                End If
            End If
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.tmrMenu_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Initialize and Load the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Load()
On Error GoTo ErrSection:

    Dim strSorting As String            ' Last known sorting value
    
    InitGrid
    LoadGrid
    
    LoadControls
    FilterGrid
    
    strSorting = GetIniFileProperty("Sorting", "", "Journals", g.strIniFile)
    If Len(strSorting) = 0 Then
        fgJournal.Col = GDCol(eGDCol_Date)
        fgJournal.Sort = flexSortGenericAscending
    Else
        fgJournal.Col = CLng(Val(Parse(strSorting, ";", 1)))
        fgJournal.Sort = CLng(Val(Parse(strSorting, ";", 2)))
    End If

    SetBackColors fgJournal
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid control
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    With fgJournal
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .HighLight = flexHighlightNever
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .SelectionMode = flexSelectionFree
        .SheetBorder = RGB(128, 128, 128)
        '.WordWrap = True
        
        .FixedCols = 0
        .Cols = GDCol(eGDCol_NumCols)
        .FixedRows = 1
        .Rows = 1
        
        .TextMatrix(0, GDCol(eGDCol_JournalID)) = "Journal ID"
        .TextMatrix(0, GDCol(eGDCol_Date)) = "Date"
        .TextMatrix(0, GDCol(eGDCol_OrderID)) = "Order ID"
        .TextMatrix(0, GDCol(eGdCol_BuySell)) = "B/S"
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_Account)) = "Account"
        .TextMatrix(0, GDCol(eGDCol_Action)) = "Action"
        .TextMatrix(0, GDCol(eGDCol_EmotionNumber)) = "Feel"
        .TextMatrix(0, GDCol(eGDCol_Feelings)) = "Feelings"
        .TextMatrix(0, GDCol(eGDCol_Reasons)) = "Reasons"
        .TextMatrix(0, GDCol(eGDCol_Thoughts)) = "Thoughts"
        .TextMatrix(0, GDCol(eGDCol_Note)) = "Notes"
        
        .ColHidden(GDCol(eGDCol_JournalID)) = True
        .ColFormat(GDCol(eGDCol_Date)) = DateFormat("Format", MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load up the grid with the journal entries
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Journal As cJournal             ' Journal object
    Dim Order As cBrokerMessage         ' Order object
    Dim OrderLeg As cBrokerMessage      ' Order Leg object
    Dim lOrderID As Long                ' Last Order ID
    Dim lAccountID As Long              ' Last Account ID
    Dim strAccount As String            ' Account name for the ID
    Dim astrAccounts As cGdArray        ' Array of unique accounts
    Dim astrSymbols As cGdArray         ' Array of unique symbols
    Dim lPos As Long                    ' Position in the array
    Dim dFromDate As Double             ' Earliest date in the journal entries
    Dim dToDate As Double               ' Latest date in the journal entries
    
    Set astrAccounts = New cGdArray
    astrAccounts.Create eGDARRAY_Strings
    
    Set astrSymbols = New cGdArray
    astrSymbols.Create eGDARRAY_Strings
    
    lOrderID = 0&
    lAccountID = 0&
    dFromDate = Date
    dToDate = 0&
    
    With fgJournal
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        g.JournalDB.LoadOrderJournals m.Journals
        
        For lIndex = 1 To m.Journals.Count
            Set Journal = m.Journals.Item(lIndex)
            If Not Journal Is Nothing Then
                If Journal.OrderID <> lOrderID Then
                    Set Order = g.AppBridge.OrderForID(Journal.OrderID)
                    
                    Set OrderLeg = New cBrokerMessage
                    OrderLeg.FromString Order("Leg1")
                End If
                
                If Order("AccountID") <> Str(lAccountID) Then
                    lAccountID = CLng(Val(Order("AccountID")))
                    strAccount = Order("AccountName")
                End If
                
                .Rows = .Rows + 1
                
                .RowData(.Rows - 1) = Order
                
                .TextMatrix(.Rows - 1, GDCol(eGDCol_JournalID)) = Journal.JournalID
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Date)) = Journal.NoteDate
                .TextMatrix(.Rows - 1, GDCol(eGDCol_OrderID)) = Order("BrokerOrderID")
                If OrderLeg("IsBuy") <> "0" Then
                    .TextMatrix(.Rows - 1, GDCol(eGdCol_BuySell)) = "Buy"
                Else
                    .TextMatrix(.Rows - 1, GDCol(eGdCol_BuySell)) = "Sell"
                End If
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Symbol)) = OrderLeg("Symbol")
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Account)) = strAccount
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Action)) = Journal.Action
                If Journal.EmotionNumber = -1 Then
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_EmotionNumber)) = ""
                Else
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_EmotionNumber)) = Str(Journal.EmotionNumber)
                End If
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Feelings)) = Journal.Feelings
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Reasons)) = Journal.WhyTrade
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Thoughts)) = Journal.Thoughts
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Note)) = Journal.Note
                
                If astrAccounts.BinarySearch(strAccount, lPos) = False Then
                    astrAccounts.Add strAccount, lPos
                End If
                
                If astrSymbols.BinarySearch(Order.Symbol, lPos) = False Then
                    astrSymbols.Add Order.Symbol, lPos
                End If
                
                If Journal.NoteDate < dFromDate Then dFromDate = Journal.NoteDate
                If Journal.NoteDate > dToDate Then dToDate = Journal.NoteDate
            End If
        Next lIndex
        
        AutoSizeGrid
        
        .Redraw = flexRDBuffered
    End With
    
    cboAccounts.Clear
    For lIndex = 0 To astrAccounts.Size - 1
        cboAccounts.AddItem astrAccounts(lIndex)
    Next lIndex
    
    cboSymbols.Clear
    For lIndex = 0 To astrSymbols.Size - 1
        cboSymbols.AddItem astrSymbols(lIndex)
    Next lIndex
    
    gdFromDate.Value = dFromDate
    If dToDate = 0# Then
        gdToDate.Value = Date
    Else
        gdToDate.Value = dToDate
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadControls
'' Description: Load up the controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadControls()
On Error GoTo ErrSection:

    Dim dTemp As Double                 ' Temporary variable

    chkAccount.Value = GetIniFileProperty("AccountOn", vbUnchecked, "Journals", g.strIniFile)
    If SetCombo(cboAccounts, GetIniFileProperty("Account", "", "Journals", g.strIniFile)) = False Then
        chkAccount.Value = vbUnchecked
    End If
    
    chkSymbol.Value = GetIniFileProperty("SymbolOn", vbUnchecked, "Journals", g.strIniFile)
    If SetCombo(cboSymbols, GetIniFileProperty("Symbol", "", "Journals", g.strIniFile)) = False Then
        chkSymbol.Value = vbUnchecked
    End If
    
    chkFromDate.Value = GetIniFileProperty("FromDateOn", vbUnchecked, "Journals", g.strIniFile)
    dTemp = GetIniFileProperty("FromDate", 0#, "Journals", g.strIniFile)
    If dTemp <> 0 Then gdFromDate.Value = dTemp
    
    chkToDate.Value = GetIniFileProperty("ToDateOn", vbUnchecked, "Journals", g.strIniFile)
    dTemp = GetIniFileProperty("ToDate", 0#, "Journals", g.strIniFile)
    If dTemp <> 0 Then gdToDate.Value = dTemp
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.LoadControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterGrid
'' Description: Filter the grid according to the filters
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterGrid()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bHide As Boolean                ' Do we want to hide the row?
    
    With fgJournal
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            bHide = False
            
            If (chkAccount.Value = vbChecked) And (cboAccounts.ListIndex >= 0) Then
                If (.TextMatrix(lIndex, GDCol(eGDCol_Account)) <> cboAccounts.Text) Then
                    bHide = True
                End If
            End If
            
            If bHide = False Then
                If (chkSymbol.Value = vbChecked) And (cboSymbols.ListIndex >= 0) Then
                    If (.TextMatrix(lIndex, GDCol(eGDCol_Symbol)) <> cboSymbols.Text) Then
                        bHide = True
                    End If
                End If
            End If
            
            If bHide = False Then
                If (chkFromDate.Value = vbChecked) Then
                    If (Val(.TextMatrix(lIndex, GDCol(eGDCol_Date))) < gdFromDate.Value) Then
                        bHide = True
                    End If
                End If
            End If
        
            If bHide = False Then
                If (chkToDate.Value = vbChecked) Then
                    If (Val(.TextMatrix(lIndex, GDCol(eGDCol_Date))) > (gdToDate.Value + 0.999999999)) Then
                        bHide = True
                    End If
                End If
            End If
            
            .RowHidden(lIndex) = bHide
        Next lIndex
        
        AutoSizeGrid
        SetBackColors fgJournal
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.FilterGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCombo
'' Description: Set the given combo box to the given value
'' Inputs:      Combo Box, Value
'' Returns:     True if Set, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetCombo(cbo As ComboBox, strValue As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    If Len(strValue) > 0 Then
        For lIndex = 0 To cbo.ListCount - 1
            If UCase(strValue) = UCase(cbo.List(lIndex)) Then
                cbo.ListIndex = lIndex
                bReturn = True
                Exit For
            End If
        Next lIndex
    End If
    
    SetCombo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.SetCombo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoSizeGrid
'' Description: Run the stuff to automatically size the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AutoSizeGrid()
On Error GoTo ErrSection:

    With fgJournal
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1, False, 75
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.AutoSizeGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteJournal
'' Description: Allow the user to delete a journal entry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteJournal()
On Error GoTo ErrSection:

    Dim lJournalID As Long              ' Journal ID from the grid
    Dim Journal As cJournal             ' Journal object

    If (fgJournal.Row >= fgJournal.FixedRows) And (fgJournal.Row < fgJournal.Rows) Then
        lJournalID = CLng(Val(fgJournal.TextMatrix(fgJournal.Row, GDCol(eGDCol_JournalID))))
        
        If InfBox("Are you sure you want to delete this journal entry?", "?", "+Yes|-No", "Journal Delete Confirmation") = "Y" Then
            fgJournal.RemoveItem fgJournal.Row
            SetBackColors fgJournal
            
            Set Journal = m.Journals.Item(Str(lJournalID))
            If Not Journal Is Nothing Then
                g.JournalDB.DeleteOrderJournal Journal
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJournals.DeleteJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveFilters
'' Description: Save the filter information to the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveFilters()
On Error GoTo ErrSection:

    SetIniFileProperty "AccountOn", chkAccount.Value, "Journals", g.strIniFile
    If chkAccount.Value = vbChecked Then
        If cboAccounts.ListIndex >= 0 Then SetIniFileProperty "Account", cboAccounts.Text, "Journals", g.strIniFile
    Else
        SetIniFileProperty "Account", "", "Journals", g.strIniFile
    End If
    
    SetIniFileProperty "SymbolOn", chkSymbol.Value, "Journals", g.strIniFile
    If chkSymbol.Value = vbChecked Then
        If cboSymbols.ListIndex >= 0 Then SetIniFileProperty "Symbol", cboSymbols.Text, "Journals", g.strIniFile
    Else
        SetIniFileProperty "Symbol", "", "Journals", g.strIniFile
    End If
    
    SetIniFileProperty "FromDateOn", chkFromDate.Value, "Journals", g.strIniFile
    If chkFromDate.Value = vbChecked Then
        SetIniFileProperty "FromDate", gdFromDate.Value, "Journals", g.strIniFile
    Else
        SetIniFileProperty "FromDate", 0#, "Journals", g.strIniFile
    End If
    
    SetIniFileProperty "ToDateOn", chkToDate.Value, "Journals", g.strIniFile
    If chkToDate.Value = vbChecked Then
        SetIniFileProperty "ToDate", gdToDate.Value, "Journals", g.strIniFile
    Else
        SetIniFileProperty "ToDate", 0#, "Journals", g.strIniFile
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmJounals.SaveFilters"
    
End Sub

