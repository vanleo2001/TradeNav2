VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmEditTicks 
   Caption         =   "Edit Ticks"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraTips 
      Height          =   1005
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   5475
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
      Caption         =   "frmEditTicks.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditTicks.frx":003A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditTicks.frx":005A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Left            =   180
         Top             =   480
         Width           =   5175
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
         Caption         =   "frmEditTicks.frx":0076
         BackColor       =   -2147483633
         ForeColor       =   8388608
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditTicks.frx":0132
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditTicks.frx":0152
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblHelpText 
         Height          =   195
         Left            =   180
         Top             =   720
         Width           =   4995
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
         Caption         =   "frmEditTicks.frx":016E
         BackColor       =   -2147483633
         ForeColor       =   8388608
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditTicks.frx":021C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditTicks.frx":023C
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   255
         Left            =   180
         Top             =   240
         Width           =   5235
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
         Caption         =   "frmEditTicks.frx":0258
         BackColor       =   -2147483633
         ForeColor       =   8388608
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmEditTicks.frx":030E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditTicks.frx":032E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   3135
      Left            =   5340
      TabIndex        =   1
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
      Caption         =   "frmEditTicks.frx":034A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditTicks.frx":0376
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditTicks.frx":0396
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL Frame1 
         Height          =   1995
         Left            =   0
         TabIndex        =   9
         Top             =   1080
         Width           =   975
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
         Caption         =   "frmEditTicks.frx":03B2
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmEditTicks.frx":03E6
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditTicks.frx":0406
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
            Height          =   345
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   735
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
            Caption         =   "frmEditTicks.frx":0422
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmEditTicks.frx":044E
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmEditTicks.frx":04B4
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdDefault 
            Height          =   345
            Left            =   120
            TabIndex        =   6
            Top             =   1560
            Width           =   735
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
            Caption         =   "frmEditTicks.frx":04D0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmEditTicks.frx":0500
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmEditTicks.frx":0584
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdGood 
            Height          =   345
            Left            =   120
            TabIndex        =   7
            Top             =   1080
            Width           =   735
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
            Caption         =   "frmEditTicks.frx":05A0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmEditTicks.frx":05CA
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmEditTicks.frx":0638
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdBad 
            Height          =   345
            Left            =   120
            TabIndex        =   8
            Top             =   660
            Width           =   735
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
            Caption         =   "frmEditTicks.frx":0654
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmEditTicks.frx":067C
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmEditTicks.frx":06E8
            RightToLeft     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   60
         TabIndex        =   3
         Top             =   480
         Width           =   855
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
         Caption         =   "frmEditTicks.frx":0704
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditTicks.frx":0732
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditTicks.frx":0752
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   375
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Width           =   855
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
         Caption         =   "frmEditTicks.frx":076E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditTicks.frx":0798
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditTicks.frx":07B8
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgTicks 
      Height          =   3435
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2955
      _cx             =   5212
      _cy             =   6059
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
   Begin HexUniControls.ctlUniLabelXP lblTicksInRed 
      Height          =   255
      Left            =   120
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
      Caption         =   "frmEditTicks.frx":07D4
      BackColor       =   -2147483633
      ForeColor       =   255
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmEditTicks.frx":0866
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditTicks.frx":0886
      RightToLeft     =   0   'False
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEditTicks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eGDCols
    eGDCol_Tick = 0
    eGDCol_Time
    eGDCol_Price
    eGDCol_Vol
    eGDCol_Default
    eGDCol_Override
    eGDCol_OldPrice
    eGDCol_OldVol
    eGDCol_TblIndexGen
    eGDCol_TblIndexUser
    eGDCol_Changed
    eGDCol_NumCols
End Enum

Private Enum eBTFields
    eBTFields_DateTime = 0              ' Double
    eBTFields_Price                     ' Double
    eBTFields_Volume                    ' Long
    eBTFields_Flags                     ' Short
    eBTFields_ScrubLevel                ' Short
    eBTFields_NewPrice                  ' Double
    eBTFields_NewVol                    ' Long
    eBTFields_Changed                   ' Short
End Enum

Private Enum eBTFlags
    eBTFlags_Scrub = 0
    eBTFlags_GenesisBad     ' delete always
    eBTFlags_GenesisRepl    ' replace
    eBTFlags_GenesisGood    ' ignore delete
    eBTFlags_UserBad
    eBTFlags_UserRepl
    eBTFlags_UserGood
End Enum
'// values returned in flags for bad tick table
'Enum eBTFlags
'{
'                                // Commands for Bad Tick Entry (stored in low-order 3 bits)
'    btDeleteLevel = 0,          //      delete if dm_badTickLevel is at or greater than entry's scrub level
'    btDeleteAlways = 1,         //      always delete ticks that match this time/price/volume
'    btReplace = 2,              //      replace with entry's new price/volume
'    btIgnoreDelete = 3,         //      ignore any deletes/replaces for this time/price/volume (usually user-supplied)
'    btAddStart = 4,             //      add a tick at start of minute with entry's new price/volume (not implemented yet)
'    btAddMiddle = 5,            //      add a tick at "middle" of minute with entry's new price/volume (not implemented yet)
'    btAddEnd = 6,               //      add a tick at end of minute with entry's new price/volume (not implemented yet)
'    btCommandMask = 0x07,       // mask for the 7 commands above
'    btTemp = 0x20,              // 0x20 => marked as temporary (snapshot) bad tick info -- never set when btUserSupplied also set
'    btUserSupplied = 0x40,      // 0x40 => user supplied this bad tick entry, 0 => Genesis supplied bad tick entry
'    btReserved = 0x80           // 0x80 => Future: use this bit carefully (preferably not at all)-- will make the flags value a negative, sorting will be affected
'};

Private Type mPrivate
    nSymbolID As Long
    nSessionDate As Long
    aEdits As cGdArray
    strPrevOverride As String
    strOldValue As String
    dtLastDataMgrTick As Double
    
    Ticks As cGdBars
    tblBadTicks As cGdTable
    aIndex As cGdArray
End Type
Private m As mPrivate

Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function

Private Function TblField(ByVal lField As eBTFields) As Long
    TblField = lField
End Function

Private Property Get BadTickFlag(ByVal lRecord As Long) As eBTFlags
On Error GoTo ErrSection:

    Dim lFlag As Long                   ' Bad tick flag from the table
    Dim lReturn As Long                 ' Return value
    
    lFlag = m.tblBadTicks(TblField(eBTFields_Flags), lRecord)
    lReturn = lFlag Mod 8
    
    BadTickFlag = lReturn
    If lReturn > 0 And GetBit(lFlag, 7) Then
        BadTickFlag = BadTickFlag + 3
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmEditTicks.BadTickFlag.Get", eGDRaiseError_Raise
    
End Property

Private Property Let BadTickFlag(ByVal lRecord As Long, ByVal Flag As eBTFlags)
On Error GoTo ErrSection:

    Dim lFlag As Long                   ' Flag to set in the table
    
    If Flag >= eBTFlags_UserBad Then
        lFlag = Flag - 3
        SetBit lFlag, 7, True
    Else
        lFlag = Flag
        SetBit lFlag, 7, False
    End If
    
    m.tblBadTicks(TblField(eBTFields_Flags), lRecord) = lFlag

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmEditTicks.BadTickFlag.Let", eGDRaiseError_Raise
    
End Property

Private Function RecordForTick(ByVal dDate#, ByVal dPrice#, ByVal lVol&, ByVal bUser As Boolean) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position of the record in the table
    Dim lChkVol As Long
    
    RecordForTick = -1&
    If lVol < 0 Then lVol = 0 '(for checking purposes)
    With m.tblBadTicks
        If .SearchAsIndex(m.aIndex, TblField(eBTFields_DateTime), dDate, lPos) Then
            For lIndex = lPos To .NumRecords - 1
                If .Item(TblField(eBTFields_DateTime), m.aIndex(lIndex)) <> dDate Then
                    Exit For
                ElseIf RoundNum(.Item(TblField(eBTFields_Price), m.aIndex(lIndex)) - dPrice, 9) = 0 Then
                    ' for checking purposes, consider a null volume to be 0
                    lChkVol = .Item(TblField(eBTFields_Volume), m.aIndex(lIndex))
                    If lChkVol < 0 Then lChkVol = 0
                    If lChkVol = lVol Then
                        If GetBit(.Item(TblField(eBTFields_Flags), m.aIndex(lIndex)), 7) = bUser Then
                            RecordForTick = m.aIndex(lIndex)
                            Exit For
                        End If
                    End If
                End If
            Next lIndex
        End If
        
        'lPos = FreeFile
        'Open "c:\chk.bin" For Binary As #lPos
        'Put #lPos, , dDate
        'Put #lPos, , .Num(0, 0)
        'Close #lPos
    End With
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditTicks.RecordForTick", eGDRaiseError_Raise
    
End Function

Private Sub cmdBad_Click()
On Error GoTo ErrSection:

    With fgTicks
        If CanEditTick(.Row) Then
            .TextMatrix(.Row, GDCol(eGDCol_Override)) = "Always Bad"
            .TextMatrix(.Row, GDCol(eGDCol_Changed)) = "1"
            ColorRow .Row
            ChangeAll .Row
        End If
    End With
    
    EnableButtons
    MoveFocus fgTicks

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.cmdBad.Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub cmdDefault_Click()
On Error GoTo ErrSection:

    Dim lRow As Long
    Dim lRec As Long                    ' Record in the table

    With fgTicks
        lRow = .Row
        If CanEditTick(lRow) Then
            .Redraw = flexRDNone
            If UCase(.TextMatrix(lRow, GDCol(eGDCol_Override))) = "EDITED" Then
                If InfBox("Changing this value will restore old values.|Do you want to continue?|", "?", "+Yes|-No", "Confirmation") = "Y" Then
                
                    lRec = CLng(.TextMatrix(lRow, GDCol(eGDCol_TblIndexGen)))
                    If lRec > -1 Then
                        If BadTickFlag(lRec) = eBTFlags_GenesisRepl Then
                            .TextMatrix(lRow, GDCol(eGDCol_Price)) = m.Ticks.PriceDisplay(m.tblBadTicks(TblField(eBTFields_NewPrice), lRec))
                            .TextMatrix(lRow, GDCol(eGDCol_Vol)) = Format(m.tblBadTicks(TblField(eBTFields_NewVol), lRec))
                            .TextMatrix(lRow, GDCol(eGDCol_Default)) = "Good"
                        Else
                            .TextMatrix(lRow, GDCol(eGDCol_Price)) = m.Ticks.PriceDisplay(m.tblBadTicks(TblField(eBTFields_Price), lRec))
                            .TextMatrix(lRow, GDCol(eGDCol_Vol)) = Format(m.tblBadTicks(TblField(eBTFields_Volume), lRec))
                            Select Case BadTickFlag(lRec)
                                Case eBTFlags_GenesisBad
                                    .TextMatrix(lRow, GDCol(eGDCol_Default)) = "Bad"
                                Case eBTFlags_GenesisGood
                                    .TextMatrix(lRow, GDCol(eGDCol_Default)) = "Good"
                                Case eBTFlags_Scrub
                                    .TextMatrix(lRow, GDCol(eGDCol_Default)) = CStr(m.tblBadTicks(TblField(eBTFields_ScrubLevel), lRec))
                                Case Else
                                    .TextMatrix(lRow, GDCol(eGDCol_Default)) = ""
                            End Select
                        End If
                    Else
                        .TextMatrix(lRow, GDCol(eGDCol_Price)) = .TextMatrix(lRow, GDCol(eGDCol_OldPrice))
                        .TextMatrix(lRow, GDCol(eGDCol_Vol)) = .TextMatrix(lRow, GDCol(eGDCol_OldVol))
                        .TextMatrix(lRow, GDCol(eGDCol_Default)) = ""
                    End If
                End If
            End If
            .TextMatrix(lRow, GDCol(eGDCol_Override)) = ""
            .TextMatrix(lRow, GDCol(eGDCol_Changed)) = "1"
            ColorRow lRow
            ChangeAll lRow
            .Redraw = True
        End If
    End With

    EnableButtons
    MoveFocus fgTicks
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.cmdDefault.Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    With fgTicks
        If CanEditTick(.Row) Then
            .Col = GDCol(eGDCol_Price)
            .EditCell
            .EditSelStart = 999
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.cmdEdit.Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub cmdGood_Click()
On Error GoTo ErrSection:

    With fgTicks
        If CanEditTick(.Row) Then
            .TextMatrix(.Row, GDCol(eGDCol_Override)) = "Always Good"
            .TextMatrix(.Row, GDCol(eGDCol_Changed)) = "1"
            ColorRow .Row
            ChangeAll .Row
        End If
    End With

    EnableButtons
    MoveFocus fgTicks

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.cmdGood.Click", eGDRaiseError_Show
    Resume ErrExit
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:
    
    If Not SaveChanges Then Exit Sub
    UpdateVisibleCharts
    
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgTicks_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim strText$, bEdited As Boolean

    With fgTicks
        If CanEditTick(Row) Then
            Select Case Col
                Case GDCol(eGDCol_Price), GDCol(eGDCol_Vol)
                    If Col = GDCol(eGDCol_Price) Then
                        strText = m.Ticks.PriceDisplay(m.Ticks.PriceFromString(.EditText))
                    Else
                        strText = Str(ValOfText(.EditText))
                    End If
                    .TextMatrix(Row, Col) = strText
                    If strText <> m.strOldValue Then
                        bEdited = True
                        '.TextMatrix(Row, GDCol(eGDCol_Default)) = "Edited"
                        .TextMatrix(Row, GDCol(eGDCol_Override)) = "Edited" '"Always Good"
                    End If
                    
                'Case GDCol(eGDCol_Override)
                '    bEdited = True
                '    .TextMatrix(Row, Col) = .EditText
            End Select
            If bEdited Then
                .TextMatrix(Row, GDCol(eGDCol_Changed)) = "1"
                ColorRow Row
                ChangeAll Row
                EnableButtons
            End If
        Else
            ' if cannot edit, then put text back to what it was
            .TextMatrix(Row, GDCol(eGDCol_Price)) = .TextMatrix(Row, GDCol(eGDCol_OldPrice))
            .TextMatrix(Row, GDCol(eGDCol_Vol)) = .TextMatrix(Row, GDCol(eGDCol_OldVol))
        End If
    End With
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.fgTicks.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgTicks_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)

    EnableButtons

End Sub

Private Sub fgTicks_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    fgTicks.ComboList = ""
    If Row < fgTicks.FixedRows Or Row >= fgTicks.Rows Then
        Cancel = True
    'ElseIf Col = GDCol(eGDCol_Override) Then
    '    m.strPrevOverride = fgTicks.TextMatrix(Row, GDCol(eGDCol_Override))
    '    fgTicks.ComboList = "Use Default|Always Good|Always Bad"
    ElseIf Col = GDCol(eGDCol_Price) Or Col = GDCol(eGDCol_Vol) Then
        m.strOldValue = fgTicks.TextMatrix(Row, Col)
        fgTicks.EditSelStart = Len(m.strOldValue) + 1
    Else
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.fgTicks.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgTicks_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GridTooltip fgTicks

End Sub

Private Sub fgTicks_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lRec As Long                    ' Record in the table
    Dim lRedraw As Long                 ' Current state of the grid's redraw

Exit Sub

    If Col = GDCol(eGDCol_Override) Then
        If fgTicks.EditText <> m.strPrevOverride And UCase(m.strPrevOverride) = "ALWAYS GOOD" Then
            If UCase(fgTicks.TextMatrix(Row, GDCol(eGDCol_Default))) = "EDITED" Then
                If InfBox("Changing this value will restore old values.|Do you want to continue?|", "?", "+Yes|-No", "Confirmation") = "Y" Then
                    With fgTicks
                        lRedraw = .Redraw
                        .Redraw = flexRDNone
                        
                        lRec = CLng(.TextMatrix(Row, GDCol(eGDCol_TblIndexGen)))
                        If lRec > -1 Then
                            If BadTickFlag(lRec) = eBTFlags_GenesisRepl Then
                                .TextMatrix(Row, GDCol(eGDCol_Price)) = m.Ticks.PriceDisplay(m.tblBadTicks(TblField(eBTFields_NewPrice), lRec))
                                .TextMatrix(Row, GDCol(eGDCol_Vol)) = Format(m.tblBadTicks(TblField(eBTFields_NewVol), lRec))
                                .TextMatrix(Row, GDCol(eGDCol_Default)) = "Good"
                            Else
                                .TextMatrix(Row, GDCol(eGDCol_Price)) = m.Ticks.PriceDisplay(m.tblBadTicks(TblField(eBTFields_Price), lRec))
                                .TextMatrix(Row, GDCol(eGDCol_Vol)) = Format(m.tblBadTicks(TblField(eBTFields_Volume), lRec))
                                Select Case BadTickFlag(lRec)
                                    Case eBTFlags_GenesisBad
                                        .TextMatrix(Row, GDCol(eGDCol_Default)) = "Bad"
                                    Case eBTFlags_GenesisGood
                                        .TextMatrix(Row, GDCol(eGDCol_Default)) = "Good"
                                    Case eBTFlags_Scrub
                                        .TextMatrix(Row, GDCol(eGDCol_Default)) = CStr(m.tblBadTicks(TblField(eBTFields_ScrubLevel), lRec))
                                    Case Else
                                        .TextMatrix(Row, GDCol(eGDCol_Default)) = ""
                                End Select
                            End If
                        Else
                            .TextMatrix(Row, GDCol(eGDCol_Price)) = .TextMatrix(Row, GDCol(eGDCol_OldPrice))
                            .TextMatrix(Row, GDCol(eGDCol_Vol)) = .TextMatrix(Row, GDCol(eGDCol_OldVol))
                            .TextMatrix(Row, GDCol(eGDCol_Default)) = ""
                        End If
                        
                        .Redraw = lRedraw
                    End With
                Else
                    Cancel = True
                End If
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.ValidateEdit", eGDRaiseError_Show
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
    RaiseError "frmEditTicks.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement of the form in the ini file
    
''    strPlacement = GetIniFileProperty("EditTicks", "", "Placement", g.strIniFile)
    If strPlacement <> "" Then
        SetFormPlacement Me, strPlacement
    Else
        CenterTheForm Me
    End If
    Icon = Picture16("kBlank")
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.Form.Load", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Public Sub ShowMe(ByVal nSymbolID As Long, Bars As cGdBars, ByVal nBar As Long)
On Error GoTo ErrSection:

    Dim bFullTicksAvail As Boolean
    Dim strMsg As String
    
    Screen.MousePointer = vbHourglass
    
    m.nSymbolID = nSymbolID
    m.nSessionDate = Bars.SessionDate(nBar)
        
    Set m.tblBadTicks = New cGdTable
    Set m.Ticks = Bars.MakeCopy(True) ' get copy of bars to duplicate properties (e.g. custom start/end)
    DM_GetBars m.Ticks, m.nSymbolID, ePRD_EachTick, m.nSessionDate, m.nSessionDate, , , , , False
    Set m.tblBadTicks = DM_GetBadTicks(m.nSymbolID, m.nSessionDate, m.nSessionDate)
    
    If IsMinutized(m.Ticks, bFullTicksAvail) Then
        If bFullTicksAvail Then
            strMsg = "The bar that you wish to edit is currently compressed data.  In order to edit this bar, you will need to download all of the ticks.||Do you wish to download now?|"
            If InfBox(strMsg, "?", "+Yes|-No", "Confirmation") = "Y" Then
                If DownloadTicks(m.Ticks, m.nSessionDate) Then
                    DM_GetBars m.Ticks, m.nSymbolID, ePRD_EachTick, m.nSessionDate, m.nSessionDate, , , , , False
                    Set m.tblBadTicks = DM_GetBadTicks(m.nSymbolID, m.nSessionDate, m.nSessionDate)
                Else
                    GoTo ErrExit
                End If
            Else
                GoTo ErrExit
            End If
        Else
            'strMsg = "Full tick data for " & DateFormat(m.nSessionDate) & " is not available for download at this time."
            strMsg = "Editing individual ticks in a bar requires the FULL TICK DATABASE (contact Genesis Sales)"
            InfBox strMsg, "I", , "Edit Ticks"
            GoTo ErrExit
        End If
    End If
    
    ' if this bar is after the DataMgr ticks, then do a splice bars to append the GenesisRT ticks
    m.dtLastDataMgrTick = m.Ticks(eBARS_DateTime, m.Ticks.Size - 1)
    If Bars(eBARS_DateTime, nBar) >= m.dtLastDataMgrTick And g.RealTime.Active Then
        g.RealTime.SpliceBars m.Ticks, , True
    End If
            
    Set m.aIndex = m.tblBadTicks.CreateIndex
    m.tblBadTicks.SortIndex m.aIndex, TblField(eBTFields_DateTime)
       
    fgTicks.Redraw = flexRDNone
    InitGrid
    LoadGrid Bars, nBar
    fgTicks.Redraw = flexRDBuffered
    
    Me.Caption = "Edit ticks for: " & SU_GetSymbol(nSymbolID)
    EnableButtons
    Screen.MousePointer = vbDefault
    ShowForm Me, True

ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmEditTicks.ShowMe", eGDRaiseError_Raise
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    If LimitFormSize(Me, fraTips.Left * 2 + fraTips.Width, _
        fraButtons.Height + fraTips.Height + 120) Then Exit Sub
    
    With fraButtons
        .Move ScaleWidth - fraButtons.Width
    End With
    
    With fraTips
        .Move .Left, Me.ScaleHeight - .Height - 60
    End With
    
    With fgTicks
        .Move .Left, .Top, fraButtons.Left - .Left * 2, fraTips.Top - .Top - 60
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "EditTicks", GetFormPlacement(Me), "Placement", g.strIniFile
    Set m.tblBadTicks = Nothing
    Set m.Ticks = Nothing
    Set m.aIndex = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw

    SetupGrid fgTicks, eGridMode_Grid
    
    With fgTicks
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Cols = GDCol(eGDCol_NumCols)
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShow
        
        .FixedRows = 1
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_Tick)) = "Tick#"
        .TextMatrix(0, GDCol(eGDCol_Time)) = "Time"
        .TextMatrix(0, GDCol(eGDCol_Price)) = "Price"
        .TextMatrix(0, GDCol(eGDCol_Vol)) = "Volume"
        .TextMatrix(0, GDCol(eGDCol_Default)) = "Default"
        .TextMatrix(0, GDCol(eGDCol_Override)) = "Override"
        .ColAlignment(GDCol(eGDCol_Tick)) = flexAlignCenterCenter
        
        .ColHidden(GDCol(eGDCol_OldPrice)) = True
        .ColHidden(GDCol(eGDCol_OldVol)) = True
        .ColHidden(GDCol(eGDCol_TblIndexGen)) = True
        .ColHidden(GDCol(eGDCol_TblIndexUser)) = True
        .ColHidden(GDCol(eGDCol_Changed)) = True

        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.InitGrid", eGDRaiseError_Raise
    
End Sub

Private Sub LoadGrid(Bars As cGdBars, ByVal nBar As Long)
On Error GoTo ErrSection:

    Dim bRC As Byte
    Dim i&, j&, dTickDate#
    Dim lRec As Long
    Dim lGenRec As Long
    Dim strDefault As String
    Dim strOverride As String
    Dim Flag As eBTFlags
    Dim dPrice As Double
    Dim lVol As Long
    Dim lLevel As Long
    Dim nGoodTicks&, nFirstTick&, nLastTick&
    Dim bBadTick As Boolean
    Dim dStartTime#, dEndTime#
    Dim strChanged As String
    Dim strText As String
    Dim aTicks As New cGdArray
            
    If Bars.Prop(eBARS_PeriodType) = ePRD_Minutes Then
        ' for minutes per bar
        dEndTime = Bars(eBARS_DateTime, nBar)
        dStartTime = Bars(eBARS_DateTime, nBar - 1)
        nFirstTick = 0
        nLastTick = m.Ticks.Size + 1000
        ' if last bar of day, move End time to the Crossover (to include trailing ticks)
        i = Int(1440 * (dEndTime - Int(dEndTime)) + 0.5)
        j = Bars.Prop(eBARS_DefaultEndTime)
        If j = 0 Then j = Bars.Prop(eBARS_EndTime)
        If i = j Then
            i = Bars.Prop(eBARS_CrossoverTime)
            If i = 0 Then i = 1440
            dEndTime = Int(dEndTime) + i / 1440#
        Else
            ' TLB 4/5/2013: or if last bar is the Suspend time, move end time to the Resume time (to include trailing ticks)
            j = Bars.Prop(eBARS_SuspendTime)
            If j > 0 And i = j Then
                i = Bars.Prop(eBARS_ResumeTime)
                dEndTime = Int(dEndTime) + i / 1440#
            End If
        End If
    Else
        ' for intraday bars other than minutes per bar
        dStartTime = 0
        dEndTime = 9999999
        ' add up the # of ticks in all previous bars of the same trading session
        nFirstTick = 1
        For i = nBar - 1 To 0 Step -1
            If Bars.SessionDate(i) <> m.nSessionDate Then Exit For
            If Bars(eBARS_DownTicks, i) >= 0 Then
                nFirstTick = nFirstTick + Bars(eBARS_UpTicks, i) + Bars(eBARS_DownTicks, i)
            Else
                nFirstTick = nFirstTick + 1
            End If
        Next
        If Bars(eBARS_DownTicks, nBar) >= 0 Then
            nLastTick = nFirstTick + Bars(eBARS_UpTicks, nBar) + Bars(eBARS_DownTicks, nBar) - 1
        Else
            nLastTick = nFirstTick
        End If
    End If
    
    If dStartTime < 0 Or dEndTime < 0 Or nFirstTick < 0 Or nLastTick < 0 Then
        InfBox "Error loading ticks for bar", "e", , "Edit Ticks"
        Exit Sub
    End If
    
    Set m.aEdits = New cGdArray
    j = 0
    nGoodTicks = 0
    fgTicks.Redraw = flexRDNone
    fgTicks.Rows = fgTicks.FixedRows
    aTicks.Size = 0
    For i = 0 To m.Ticks.Size - 1
        dTickDate = m.Ticks(eBARS_DateTime, i)
        If dTickDate >= dEndTime Then
            Exit For
        ElseIf dTickDate >= dStartTime Then
            strDefault = ""
            dPrice = m.Ticks(eBARS_Close, i)
            lVol = m.Ticks(eBARS_Vol, i)
            lRec = RecordForTick(dTickDate, m.Ticks(eBARS_Close, i), m.Ticks(eBARS_Vol, i), False)
            If lRec > -1 Then
                Flag = BadTickFlag(lRec)
                Select Case Flag
                    Case eBTFlags_Scrub
                        strDefault = CStr(m.tblBadTicks(TblField(eBTFields_ScrubLevel), lRec))
                    Case eBTFlags_GenesisBad
                        strDefault = "Bad"
                    Case eBTFlags_GenesisRepl
                        strDefault = "Good"
                        dPrice = m.tblBadTicks(TblField(eBTFields_NewPrice), lRec)
                        lVol = m.tblBadTicks(TblField(eBTFields_NewVol), lRec)
                    Case eBTFlags_GenesisGood
                        strDefault = "Good"
                End Select
            End If
            lGenRec = lRec
                
            strOverride = "" ' "Use Default"
            lRec = RecordForTick(dTickDate, m.Ticks(eBARS_Close, i), m.Ticks(eBARS_Vol, i), True)
            If lRec > -1 Then
                Flag = BadTickFlag(lRec)
                
                Select Case Flag
                    Case eBTFlags_UserBad
                        strOverride = "Always Bad"
                    Case eBTFlags_UserRepl
                        ''strDefault = "Edited"
                        strOverride = "Edited" '"Always Good"
                        dPrice = m.tblBadTicks(TblField(eBTFields_NewPrice), lRec)
                        lVol = m.tblBadTicks(TblField(eBTFields_NewVol), lRec)
                    Case eBTFlags_UserGood
                        strOverride = "Always Good"
                End Select
            End If
            
            bBadTick = False
            If Len(strOverride) = 0 Then ' UCase(strOverride) = "USE DEFAULT" Then
                lLevel = CLng(ValOfText(strDefault))
                If lLevel > 0 Then
                    If lLevel <= g.iScrubLevel Then bBadTick = True
                ElseIf UCase(strDefault) = "BAD" Then
                    bBadTick = True
                End If
            ElseIf UCase(strOverride) = "ALWAYS BAD" Then
                bBadTick = True
            End If
            If bBadTick = False Then
                nGoodTicks = nGoodTicks + 1
            End If
            
            ' wait to exit for loop until beyond # good ticks (to show any ending bad ticks)
            If nGoodTicks > nLastTick Then Exit For
            
            ' show if at # good ticks or the bad tick just prior
            If nGoodTicks >= nFirstTick Or (nGoodTicks = nFirstTick - 1 And bBadTick) Then
                j = j + 1
                If dTickDate > m.dtLastDataMgrTick Then
                    strChanged = "-99"  ' cannot edit GenesisRT ticks
                Else
                    strChanged = "0"
                End If
                strText = Str(j) & vbTab & Format(dTickDate, "hh:mm:ss") _
                    & vbTab & m.Ticks.PriceDisplay(dPrice) & vbTab & Format(lVol) _
                    & vbTab & strDefault & vbTab & strOverride _
                    & vbTab & m.Ticks.PriceDisplay(m.Ticks(eBARS_Close, i)) _
                    & vbTab & Format(m.Ticks(eBARS_Vol, i)) & vbTab & Str(lGenRec) _
                    & vbTab & Str(lRec) & vbTab & strChanged
                fgTicks.AddItem strText
                If IsIDE Then
                    aTicks.Add strText
                End If
                    
                ColorRow fgTicks.Rows - 1
            End If
        End If
    Next
    fgTicks.AutoSize 0, fgTicks.Cols - 1, False, 75
    fgTicks.Redraw = flexRDBuffered

    'need to determine whether data is minute-ized or all ticks
    'if minute-ized then cannot do network or volume editing
    If m.Ticks.IsActiveArray(eBARS_Vol) Then
        fgTicks.ColHidden(GDCol(eGDCol_Vol)) = False
        lblHelpText.Caption = "* To edit a price or volume, double-click the cell."
    Else
        fgTicks.ColHidden(GDCol(eGDCol_Vol)) = True
        lblHelpText.Caption = "* To edit a price, double-click the cell."
    End If
    
    If IsIDE Then
        aTicks.ToFile App.Path & "\Chk\Ticks.txt"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.LoadGrid", eGDRaiseError_Raise

End Sub

Private Sub ColorRow(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim lLevel As Long                  ' Scrub level for the tick
    Dim strDefault As String            ' Current Genesis Default for the tick
    Dim strOverride As String           ' Current User Override for the tick
    Dim bColor As Boolean               ' Color the tick bad?
    
    bColor = False
    strDefault = fgTicks.TextMatrix(lRow, GDCol(eGDCol_Default))
    strOverride = fgTicks.TextMatrix(lRow, GDCol(eGDCol_Override))
    
    If Len(strOverride) = 0 Then ' UCase(strOverride) = "USE DEFAULT" Then
        lLevel = CLng(ValOfText(strDefault))
        If lLevel > 0 Then
            If lLevel <= g.iScrubLevel Then bColor = True
        ElseIf UCase(strDefault) = "BAD" Then
            bColor = True
        End If
    ElseIf UCase(strOverride) = "ALWAYS BAD" Then
        bColor = True
    End If
    
    If bColor Then
        fgTicks.Cell(flexcpForeColor, lRow, 0, lRow, fgTicks.Cols - 1) = vbRed
    Else
        fgTicks.Cell(flexcpForeColor, lRow, 0, lRow, fgTicks.Cols - 1) = vbBlack
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.ColorRow", eGDRaiseError_Raise
    
End Sub

Private Sub ChangeAll(ByVal lSourceRow As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim dGridTime As Double             ' Time of the current row in the grid
    Dim lCol As Long                    ' Index into a for loop
    Dim dTime As Double                 ' Time on the source row
    Dim dPrice As Double                ' Original Price on the source row
    Dim lVol As Long                    ' Original Volume on the source row
    
    With fgTicks
        lRedraw = .Redraw
        .Redraw = flexRDNone
            
        ' Get the "key" from the source row...
        dTime = TimeFromString(.TextMatrix(lSourceRow, GDCol(eGDCol_Time)))
        dPrice = m.Ticks.PriceFromString(.TextMatrix(lSourceRow, GDCol(eGDCol_OldPrice)))
        lVol = CLng(ValOfText(.TextMatrix(lSourceRow, GDCol(eGDCol_OldVol))))
    
        ' Walk through all of the rows in the grid...
        For lIndex = .FixedRows To .Rows - 1
            ' Get the time off of the current row...
            dGridTime = TimeFromString(.TextMatrix(lIndex, GDCol(eGDCol_Time)))
            
            ' If the time is greater than the time on the source row, stop...
            If dGridTime > dTime Then
                Exit For
                
            ' Otherwise if the time matches and the price and volume match, change
            ' the row...
            ElseIf dGridTime = dTime And lIndex <> lSourceRow Then
                If m.Ticks.PriceFromString(.TextMatrix(lIndex, GDCol(eGDCol_OldPrice))) = dPrice And ValOfText(.TextMatrix(lIndex, GDCol(eGDCol_OldVol))) = lVol Then
                    For lCol = 0 To .Cols - 1
                        .TextMatrix(lIndex, lCol) = .TextMatrix(lSourceRow, lCol)
                    Next lCol
                    
                    ColorRow lIndex
                End If
            End If
        Next lIndex
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditTicks.ChangeAll", eGDRaiseError_Raise
    
End Sub

Private Function TimeFromString(ByVal strTime As String) As Double
On Error GoTo ErrSection:

    Dim dTime As Double                 ' Time to return
    
    dTime = (Val(Parse(strTime, ":", 1)) / 24#) + (Val(Parse(strTime, ":", 2)) / 1440#) + (Val(Parse(strTime, ":", 3)) / 86400#)
    TimeFromString = dTime

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditTicks.TimeFromString", eGDRaiseError_Raise
    
End Function

Private Function SaveChanges() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim tblChanges As New cGdTable
    Dim lRec As Long
    Dim dTime As Double
    Dim lUserRec As Long
    Dim dPrice As Double
    Dim lVol As Long
    Dim dNewPrice As Double
    Dim lNewVol As Long
    Dim lUserFlag As Long
    Dim strFixFile As String
    Dim strText As String
    Dim aStrings As New cGdArray
    
    ' see if this is a Genesis user (allowed to send edits to the Master database)
    If FileExist("c:\common\ask.exe") And g.lLCD > 0 Then
        aStrings.FromFile "k:\common\FixTicks.dir"
        If Len(aStrings(0)) > 0 Then
            ' see if they are authorized (have write-access)
            strFixFile = AddSlash(aStrings(0)) & Str(g.lLCD)
            If Not aStrings.ToFile(strFixFile & ".TMP") Then
                strFixFile = ""
            Else
                KillFile strFixFile & ".TMP"
                strFixFile = strFixFile & ".FIX"
                strText = "Also save changes to the Master database| for all Genesis customers?"
                Select Case InfBox(strText, "?", "+Local|Master|-Cancel", "Save changes to ...")
                Case "C"
                    Exit Function
                'Case "M"
                    'strFixFile = AddSlash(aStrings(0)) & Str(g.lLCD)
                    ' make sure they are authorized (have write-access)
                    'If Not aStrings.ToFile(strFixFile & ".TMP") Then
                    '    Beep
                    '    InfBox "You are not currently authorized to save changes to the master database.", "e", , "Permission Denied"
                    '    Exit Function
                    'End If
                    'KillFile strFixFile & ".TMP"
                    'strFixFile = strFixFile & ".FIX"
                Case "L"
                    strFixFile = ""
                End Select
            End If
        End If
    End If
    
    lUserFlag = 64 ' User flag (when user adding to their own database)
    If Len(strFixFile) > 0 Then
        'TLB: the Temp flag doesn't really work since the user can't undo a Bad tick sent
        'to the master if they change their mind.
        ''lUserFlag = 32 ' Temp flag (when Genesis person adding to master database)
    End If

    Set tblChanges = m.tblBadTicks.MakeCopy
    tblChanges.NumRecords = 0
    tblChanges.CreateField eGDARRAY_Shorts, , "Changed"
    
    With fgTicks
        lRec = 0&
        aStrings.Clear
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, GDCol(eGDCol_Changed)) = "1" Then
                ''If Not ((UCase(.TextMatrix(lIndex, GDCol(eGDCol_Override))) = "USE DEFAULT") And (.TextMatrix(lIndex, GDCol(eGDCol_TblIndexUser)) = "-1")) Then
                If Len(.TextMatrix(lIndex, GDCol(eGDCol_Override))) > 0 Or .TextMatrix(lIndex, GDCol(eGDCol_TblIndexUser)) <> "-1" Then
                    dTime = m.nSessionDate + TimeFromString(.TextMatrix(lIndex, GDCol(eGDCol_Time)))
                    dPrice = m.Ticks.PriceFromString(.TextMatrix(lIndex, GDCol(eGDCol_OldPrice)))
                    lVol = CLng(ValOfText(.TextMatrix(lIndex, GDCol(eGDCol_OldVol))))
                    If lVol < 0 Then lVol = 0
                    
                    If Not RecExists(tblChanges, dTime, dPrice, lVol) Then
                        tblChanges(TblField(eBTFields_DateTime), lRec) = dTime
                        tblChanges(TblField(eBTFields_Price), lRec) = dPrice
                        tblChanges(TblField(eBTFields_Volume), lRec) = lVol
                        tblChanges(TblField(eBTFields_ScrubLevel), lRec) = 0
                        dNewPrice = m.Ticks.PriceFromString(.TextMatrix(lIndex, GDCol(eGDCol_Price)))
                        lNewVol = CLng(ValOfText(.TextMatrix(lIndex, GDCol(eGDCol_Vol))))
                        tblChanges(TblField(eBTFields_NewPrice), lRec) = dNewPrice
                        tblChanges(TblField(eBTFields_NewVol), lRec) = lNewVol
                        lUserRec = CLng(ValOfText(.TextMatrix(lIndex, GDCol(eGDCol_TblIndexUser))))
                        
                        Select Case UCase(.TextMatrix(lIndex, GDCol(eGDCol_Override)))
                            Case "EDITED"
                                tblChanges(TblField(eBTFields_Flags), lRec) = (eBTFlags_UserRepl - 3) + lUserFlag
                                strText = "Edited"
                                If lUserRec = -1& Then
                                    tblChanges(TblField(eBTFields_Changed), lRec) = -1
                                Else
                                    tblChanges(TblField(eBTFields_Changed), lRec) = m.tblBadTicks(TblField(eBTFields_Flags), lUserRec)
                                End If
                            
                            Case "ALWAYS GOOD"
                                'If UCase(.TextMatrix(lIndex, GDCol(eGDCol_Default))) = "EDITED" Then
                                '    tblChanges(TblField(eBTFields_Flags), lRec) = (eBTFlags_UserRepl - 3) + lUserFlag
                                '    strText = "Edited"
                                'Else
                                    tblChanges(TblField(eBTFields_Flags), lRec) = (eBTFlags_UserGood - 3) + lUserFlag
                                    strText = "Always Good"
                                'End If
                                If lUserRec = -1& Then
                                    tblChanges(TblField(eBTFields_Changed), lRec) = -1
                                Else
                                    tblChanges(TblField(eBTFields_Changed), lRec) = m.tblBadTicks(TblField(eBTFields_Flags), lUserRec)
                                End If
                            
                            Case "ALWAYS BAD"
                                tblChanges(TblField(eBTFields_Flags), lRec) = (eBTFlags_UserBad - 3) + lUserFlag
                                If lUserRec = -1& Then
                                    tblChanges(TblField(eBTFields_Changed), lRec) = -1
                                Else
                                    tblChanges(TblField(eBTFields_Changed), lRec) = m.tblBadTicks(TblField(eBTFields_Flags), lUserRec)
                                End If
                                strText = "Always Bad"
        
                            Case Else
                                tblChanges(TblField(eBTFields_Flags), lRec) = m.tblBadTicks(TblField(eBTFields_Flags), lUserRec)
                                If lUserRec <> -1& Then
                                    tblChanges(TblField(eBTFields_Changed), lRec) = -2
                                End If
                                strText = "Default"
                        End Select
                        
                        If Len(strFixFile) > 0 Then
                            strText = Str(g.lLCD) & vbTab & Environ("UserName") & vbTab & GetSymbol(m.nSymbolID) & vbTab & Str(m.nSymbolID) _
                                & vbTab & Format(dTime, "yyyymmdd") & vbTab & Format(dTime, "HhNn:Ss") & vbTab _
                                & Str(dPrice) & vbTab & Str(lVol) & vbTab & strText & vbTab & Str(dNewPrice) & vbTab & Str(lNewVol)
                            aStrings.Add strText
                        End If
                        
                        lRec = lRec + 1
                    End If
                End If
            End If
        Next lIndex
    End With
    
    ' save changes to local database
    DM_PutBadTicks m.nSymbolID, tblChanges
    
    ' save to master (only Genesis employees)
    If aStrings.Size > 0 Then
        If aStrings.ToFile(strFixFile, True) Then
            strText = Str(aStrings.Size) & " change(s) will be saved to the |Master database and downloaded later."
            InfBox strText, "i", , "Saved to Master database"
        Else
            strText = "The changes were NOT saved to the master database|(you may not be authorized for this)."
            Beep
            InfBox strText, "e", , "Error saving to Master database"
        End If
    End If

'FileFromString "DAJ.TXT", tblChanges.ToString(vbCrLf, vbTab)

    SaveChanges = True

ErrExit:
    Set tblChanges = Nothing
    Exit Function
    
ErrSection:
    Set tblChanges = Nothing
    RaiseError "frmEditTicks.SaveChanges", eGDRaiseError_Raise

End Function

Private Function RecExists(tbl As cGdTable, ByVal dDate#, ByVal dPrice#, ByVal lVol&) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    
    RecExists = False
    For lIndex = 0 To tbl.NumRecords - 1
        If tbl(TblField(eBTFields_DateTime), lIndex) = dDate Then
            If tbl(TblField(eBTFields_Price), lIndex) = dPrice Then
                If tbl(TblField(eBTFields_Volume), lIndex) = lVol Then
                    RecExists = True
                    Exit For
                End If
            End If
        End If
    Next lIndex

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditTicks.RecExists", eGDRaiseError_Raise
    
End Function

Private Sub EnableButtons()

    On Error Resume Next
    With fgTicks
        If .Row >= .FixedRows And .Row < .Rows Then
            Select Case UCase(.TextMatrix(.Row, GDCol(eGDCol_Override)))
            Case "ALWAYS GOOD", "EDITED"
                Enable cmdBad
                Disable cmdGood
                Enable cmdDefault
            Case "ALWAYS BAD"
                Disable cmdBad
                Enable cmdGood
                Enable cmdDefault
            Case Else
                Enable cmdBad
                Enable cmdGood
                Disable cmdDefault
            End Select
            Enable cmdEdit
        Else
            Disable cmdBad
            Disable cmdGood
            Disable cmdDefault
            Disable cmdEdit
        End If
    End With

End Sub

Private Function CanEditTick(ByVal nRow&) As Boolean
On Error GoTo ErrSection:

    With fgTicks
        If nRow >= .FixedRows And nRow < .Rows Then
            If .TextMatrix(nRow, GDCol(eGDCol_Changed)) = "-99" Then
                InfBox "This tick came after the last quote board refresh.|You will need to do another quote board refresh before this tick can be edited.", "!", , "Edit Tick"
            Else
                CanEditTick = True
            End If
        End If
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditTicks.CanEditTick"
End Function

