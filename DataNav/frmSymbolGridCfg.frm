VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmSymbolGridCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraView 
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   3420
      Width           =   4275
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
      Caption         =   "frmSymbolGridCfg.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSymbolGridCfg.frx":0028
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSymbolGridCfg.frx":0048
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniRadioXP optSector 
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1695
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
         Caption         =   "frmSymbolGridCfg.frx":0064
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmSymbolGridCfg.frx":00A8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGridCfg.frx":00C8
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optListView 
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   300
         Width           =   675
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
         Caption         =   "frmSymbolGridCfg.frx":00E4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmSymbolGridCfg.frx":0110
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGridCfg.frx":0130
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin MSComctlLib.ImageCombo cboFilters 
         Height          =   330
         Left            =   960
         TabIndex        =   12
         Top             =   255
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "ImageCombo1"
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraFields 
      Height          =   1995
      Left            =   120
      TabIndex        =   3
      Top             =   1260
      Width           =   4275
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
      Caption         =   "frmSymbolGridCfg.frx":014C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSymbolGridCfg.frx":0178
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSymbolGridCfg.frx":0198
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkShowFlags 
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1560
         Width           =   1635
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
         Caption         =   "frmSymbolGridCfg.frx":01B4
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmSymbolGridCfg.frx":01F4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGridCfg.frx":0214
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optDescending 
         Height          =   255
         Left            =   1740
         TabIndex        =   8
         Top             =   1200
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
         Caption         =   "frmSymbolGridCfg.frx":0230
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmSymbolGridCfg.frx":0266
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGridCfg.frx":0286
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optAscending 
         Height          =   255
         Left            =   540
         TabIndex        =   7
         Top             =   1200
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
         Caption         =   "frmSymbolGridCfg.frx":02A2
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmSymbolGridCfg.frx":02D6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGridCfg.frx":02F6
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboFields 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   840
         Width           =   2595
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
         Tip             =   "frmSymbolGridCfg.frx":0312
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGridCfg.frx":0332
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdFields 
         Height          =   435
         Left            =   180
         TabIndex        =   4
         Top             =   300
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
         Caption         =   "frmSymbolGridCfg.frx":034E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSymbolGridCfg.frx":038A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGridCfg.frx":03AA
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblRank 
         Height          =   195
         Left            =   180
         Top             =   900
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
         Caption         =   "frmSymbolGridCfg.frx":03C6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolGridCfg.frx":0406
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGridCfg.frx":0426
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   975
      Left            =   4560
      TabIndex        =   14
      Top             =   120
      Width           =   1035
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
      Caption         =   "frmSymbolGridCfg.frx":0442
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSymbolGridCfg.frx":046E
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSymbolGridCfg.frx":048E
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   435
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1035
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
         Caption         =   "frmSymbolGridCfg.frx":04AA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSymbolGridCfg.frx":04D0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGridCfg.frx":04F0
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   5
         Top             =   540
         Width           =   1035
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
         Caption         =   "frmSymbolGridCfg.frx":050C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSymbolGridCfg.frx":053A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGridCfg.frx":055A
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraFonts 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4275
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
      Caption         =   "frmSymbolGridCfg.frx":0576
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSymbolGridCfg.frx":05A0
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSymbolGridCfg.frx":05C0
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdGridFont 
         Height          =   435
         Left            =   180
         TabIndex        =   1
         Top             =   300
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
         Caption         =   "frmSymbolGridCfg.frx":05DC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSymbolGridCfg.frx":0610
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGridCfg.frx":0630
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblGridFontSample 
         Height          =   495
         Left            =   1440
         Top             =   300
         Width           =   2715
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
         Caption         =   "frmSymbolGridCfg.frx":064C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmSymbolGridCfg.frx":068C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSymbolGridCfg.frx":06AC
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmSymbolGridCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmSymbolGridCfg.frm
'' Description: Allows the user to set some configuration settings for the
''              symbol grid form
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?
    strDisplayFields As String          ' Fields displayed in the grid
    strRankField As String              ' Field the user wants to rank
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(strGridFont As String, strDisplayFields As String, strRankField As String, _
                        bAscending As Boolean, bShowFlags As Boolean, bList As Boolean, _
                        strFilterID As String) As Boolean
On Error GoTo ErrSection:

    ' Set up the font information...
    FontFromString lblGridFontSample.Font, strGridFont
    
    ' Set up the fields information...
    m.strDisplayFields = strDisplayFields
    m.strRankField = strRankField
    LoadFieldsCombo
    If bAscending = True Then
        optAscending.Value = True
        optDescending.Value = False
    Else
        optAscending.Value = False
        optDescending.Value = True
    End If
    If bShowFlags Then
        chkShowFlags.Value = vbChecked
    Else
        chkShowFlags.Value = vbUnchecked
    End If
    
    ' Set up the view information...
    If bList Then
        optListView.Value = True
        optSector.Value = False
    Else
        optListView.Value = False
        optSector.Value = True
    End If
    cboFilters.ImageList = frmMain.img16
    LoadFiltersCombo strFilterID
    cboFilters.Locked = True

    EnableControls
    ShowForm Me, True
    ShowMe = m.bOK
    
    If m.bOK Then
        strGridFont = FontToString(lblGridFontSample.Font)
        strDisplayFields = m.strDisplayFields
        strRankField = cboFields.Text
        bAscending = optAscending.Value
        bShowFlags = (chkShowFlags.Value = vbChecked)
        bList = optListView.Value
        strFilterID = cboFilters.SelectedItem.Key
    End If

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmSymbolGridCfg.ShowMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Hide the form and let ShowMe know not to save information
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
    RaiseError "frmSymbolGridCfg.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdFields_Click
'' Description: Allow the user to change what fields are showing
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdFields_Click()
On Error GoTo ErrSection:

    Dim strIgnore As String             ' List of fields not to show in lists
    Dim bScansOn As Boolean             ' Are criteria and filters turned on?
    Dim astrAvailable As New cGdArray   ' Array of available fields
    Dim astrUsed As New cGdArray        ' Array of used fields
    Dim astrUsedSorted As New cGdArray  ' Array of used fields sorted
    Dim astrFields As New cGdArray      ' Array of fields used
    Dim lIndex As Long                  ' Index into a for loop
    Dim strID As String                 ' ID of the item in the symbol pool
    Dim lField As Long                  ' Field number for item in the pool
    Dim strName As String               ' Name of the symbol pool item
    Dim obj As Object                   ' Object from the symbol pool
    Dim bSkip As Boolean                ' Should we skip this one?

    strIgnore = "||FLAGS|FLAGGED SYMBOLS|SYMINDEX|SYMBOL|DBRECNUM|ALL SYMBOLS|"
    
    bScansOn = ScansEnabled

    astrAvailable.Create eGDARRAY_Strings
    astrUsed.Create eGDARRAY_Strings
    astrUsedSorted.Create eGDARRAY_Strings

    astrFields.SplitFields m.strDisplayFields, "|"
    For lIndex = 0 To astrFields.Size - 1
        strID = Parse(astrFields(lIndex), "\", 1)
        lField = g.SymbolPool.FieldNumForID(strID)
        If lField >= 0 Then
            strName = g.SymbolPool.ArrayTable.FieldName(lField)
            'skip certain fields
            If InStr(strIgnore, "|" & UCase(strName) & "|") = 0 Then
                astrUsed.Add strName
                astrUsedSorted.Add strName
            End If
        End If
    Next
    astrUsedSorted.Sort
    
    For lIndex = 0 To g.SymbolPool.ArrayTable.NumFields - 1
        strID = g.SymbolPool.FieldID(lIndex)
        If Len(strID) > 0 And Left(strID, 4) <> "DSP:" Then
            Set obj = g.SymbolPool.PoolObject(strID)
            
            If bScansOn = False And (Left(strID, 4) = "DSV:" Or Left(strID, 4) = "FIL:") Then
                bSkip = True
            Else
                bSkip = False
            End If
            
            If Not obj Is Nothing Then
                If bSkip = False Then bSkip = (obj.IsActive <> True)
            End If
            
            strName = g.SymbolPool.ArrayTable.FieldName(lIndex)
            'skip certain fields
            If InStr(strIgnore, "|" & UCase(strName) & "|") = 0 Then
                If astrUsedSorted.BinarySearch(strName) = False And bSkip = False Then
                    astrAvailable.Add strName
                End If
            End If
        End If
    Next lIndex
    astrAvailable.Sort eGdSort_IgnoreCase

    ' Call the add/remove form
    If frmAddRemove.ShowMe(astrAvailable, astrUsed, eOrderMode_Ordered, , "Fields to Display in Symbol Grid") = True Then
        m.strDisplayFields = "GRP:_FLAGS_.GRP|INF:SYMBOL"
        For lIndex = 0 To astrUsed.Size - 1
            lField = g.SymbolPool.ArrayTable.FieldNum(astrUsed(lIndex))
            strID = g.SymbolPool.FieldID(lField)
            If Len(strID) > 0 Then
                m.strDisplayFields = m.strDisplayFields & "|" & strID
            End If
        Next lIndex
        
        LoadFieldsCombo
    End If
    
ErrExit:
    Set astrFields = Nothing
    Set astrAvailable = Nothing
    Set astrUsed = Nothing
    Set astrUsedSorted = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGridCfg.cmdFields.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdGridFont_Click
'' Description: Allow the user to change the font
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdGridFont_Click()
On Error GoTo ErrSection:
    
    CommonDialogFont frmMain.CommonDialog1, lblGridFontSample.Font

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGridCfg.cmdGridFont.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Hide the form and let the ShowMe know to save information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGridCfg.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Caption = "Symbol Grid Settings"
    Icon = Picture16(ToolbarIcon("ID_SymbolGrid"))
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGridCfg.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Hide the form and let ShowMe know not to save information
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGridCfg.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFiltersCombo
'' Description: Load the filters combo box with the appropriate items
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFiltersCombo(ByVal strSelID As String)
On Error Resume Next

    Dim strKey$
    Dim bScansOn As Boolean             ' Are criteria and filters turned on?
    Dim lIndex As Long                  ' Index into a for loop
    Dim strID As String                 ' ID for item in the symbol pool
    Dim strType As String               ' Type of the item in the symbol pool
    Dim strPicture As String            ' Picture to show in the combo box
    Dim obj As Object                   ' Object from the symbol pool
    Dim lNode As Long                   ' Index into a collection in the pool
    Dim bSelExists As Boolean           ' Does the selection exist?
    Dim lSortStart As Long              ' Where to start the sort
    Dim lAt As Long                     ' Where are we at?
    Dim astrItems As New cGdArray       ' Items to add to the combo box
    Dim strItem As String               ' Item to add to the combo box

    bScansOn = ScansEnabled
        
    If cboFilters.ComboItems.Count > 0 Then
        strSelID = cboFilters.SelectedItem.Key
        cboFilters.ComboItems.Clear
    End If
    
    With g.SymbolPool
        For lIndex = 0 To .ArrayTable.NumFields - 1
            strID = .FieldID(lIndex)
            If Len(strID) = 0 Then
                strType = "" '???
            Else
                strType = Left(strID, 3)
                strPicture = ""
                Set obj = .PoolObject(strID)
                Select Case strType
                    Case "GRP":
                        strPicture = ToolbarIcon("ID_SymbolGroups")
                        If Len(strSelID) = 0 Then
                            If UCase(obj.Name) = "HUME" Then
                                strSelID = .FieldID(lIndex)
                            End If
                        End If
                    
                    Case "FIL":
                        If bScansOn Then strPicture = ToolbarIcon("ID_Filters")
                    
                    Case "DSV":
                        If bScansOn Then
                            'only if boolean
                            strKey = Mid(strID, 5)
                            lNode = .Criterias.Index(strKey)
                            If lNode > 0 Then
                                If .Criterias(lNode).IsBoolean Then
                                    strPicture = ToolbarIcon("ID_Criteria")
                                End If
                            End If
                        End If
                End Select
                
                If Len(strPicture) > 0 Then
                    If obj.IsActive = True Then
                        If strID = strSelID Then
                            bSelExists = True
                        End If
                        
                        If lSortStart = 0 And lIndex >= g.SymbolPool.OtherFieldsStart Then
                            lSortStart = astrItems.Size
                        End If
                        
                        ' keep "flagged symbols" above where we sort
                        lAt = -1
                        If strID = "GRP:_FLAGS_.GRP" And lSortStart > 0 Then
                            lAt = lSortStart
                            lSortStart = lSortStart + 1
                        End If
                        
                        astrItems.Add .ArrayTable.FieldName(lIndex) & vbTab _
                                & strID & vbTab & strPicture, lAt
                    End If
                End If
            End If
        Next
    End With
    
    If lSortStart > 0 Then
        astrItems.Sort eGdSort_IgnoreCase, lSortStart
    End If

    For lIndex = 0 To astrItems.Size - 1
        strItem = astrItems(lIndex)
        cboFilters.ComboItems.Add , Parse(strItem, vbTab, 2), _
            Parse(strItem, vbTab, 1), Parse(strItem, vbTab, 3)
    Next


    If bSelExists Then
        cboFilters.ComboItems(strSelID).Selected = True
    Else
        cboFilters.ComboItems(1).Selected = True
    End If

    cboFilters.Refresh

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFieldsCombo
'' Description: Load the fields combo box with the used fields
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFieldsCombo()
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray      ' Array of used fields
    Dim lIndex As Long                  ' Index into a for loop
    Dim strID As String                 ' ID of the item in the symbol pool
    Dim lField As Long                  ' Field number for item in the pool
    Dim lSelect As Long                 ' Item to select
    
    astrFields.SplitFields m.strDisplayFields, "|"
    
    With cboFields
        .Clear
        lSelect = -1&
        For lIndex = 0 To astrFields.Size - 1
            strID = Parse(astrFields(lIndex), "\", 1)
            lField = g.SymbolPool.FieldNumForID(strID)
            
            .AddItem g.SymbolPool.ArrayTable.FieldName(lField)
            If .List(.NewIndex) = m.strRankField Then
                lSelect = .NewIndex
            End If
        Next lIndex
        
        If lSelect = -1& Then
            .ListIndex = 1
        Else
            .ListIndex = lSelect
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGridCfg.LoadFieldsCombo", eGDRaiseError_Raise
    
End Sub

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

    Enable cboFilters, optListView.Value

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGridCfg.EnableControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optListView_Click
'' Description: Enable/Disable the combo box depending on value of list option
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optListView_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGridCfg.optListView.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    optSector_Click
'' Description: Enable/Disable the combo box depending on value of list option
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optSector_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSymbolGridCfg.optSector.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub


