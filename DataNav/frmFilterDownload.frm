VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmFilterDownload 
   Caption         =   "Title"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   400
      Left            =   60
      TabIndex        =   1
      Top             =   6060
      Width           =   5775
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
      Caption         =   "frmFilterDownload.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmFilterDownload.frx":0034
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmFilterDownload.frx":0054
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   4530
         TabIndex        =   3
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
         Caption         =   "frmFilterDownload.frx":0070
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmFilterDownload.frx":009C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmFilterDownload.frx":00BC
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   3240
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
         Caption         =   "frmFilterDownload.frx":00D8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmFilterDownload.frx":00FC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmFilterDownload.frx":011C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblDownload 
         Height          =   315
         Left            =   60
         Top             =   60
         Width           =   3015
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
         Caption         =   "frmFilterDownload.frx":0138
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmFilterDownload.frx":0190
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmFilterDownload.frx":01B0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgDownload 
      Height          =   675
      Left            =   900
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
      _cx             =   2355
      _cy             =   1191
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
   Begin HexUniControls.ctlUniLabelXP lblGeneralInfo 
      Height          =   375
      Left            =   60
      Top             =   60
      Width           =   5775
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
      Caption         =   "frmFilterDownload.frx":01CC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmFilterDownload.frx":01EC
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmFilterDownload.frx":020C
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmFilterDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const kSupportMsg = "Please contact Genesis support to enable or disable this feature."

'These fields correspond to the columns in the config file and grid
Private Enum eCfgFields
    eCfgTypeField = 0
    eCfgDescField
    eCfgDailyField
    eCfgIntradayField
    eCfgToolTipField
End Enum

'These fields correspond to files in table
Private Enum eTableFields
    eTableAuthStr = 0
    eTableDownloadStr
    eTableDownloadDefault
    eTableDownloadSize
    eTableExclude
    eTableInclude
    eTableHasModule
    eTableIgnoreSym
End Enum

Private Type mPrivate
    aCfgFile As New cGdArray
    tbDownloadInfo As cGdTable
    aIdxByDwnloadStr As cGdArray
    nMaxCol As Long
    bHideCancel As Boolean
    strSupportMsg As String
    strSizeLabel As String
    strSize As String
    strKbMb As String
    bExtremeChartsMode As Boolean
End Type

Private m As mPrivate


Public Sub ShowMe(Optional ByVal bHiddenSave As Boolean = False)
On Error GoTo ErrSection:

    Dim strFile$
    
    strFile = g.strAppPath & "\Provided\DownloadFilter.cfg"
    m.aCfgFile.FromFile strFile
    m.strSizeLabel = "Avg Download Size ="
    
    If m.aCfgFile.Size = 0 Then
        Unload Me
        Exit Sub
    End If
    
    If ExtremeCharts = 1 And Not HasModule("F") And Not HasModule("IT") And Not HasModule("ST") Then
        m.bExtremeChartsMode = True
        Me.Height = 4000
    Else
        m.bExtremeChartsMode = False
    End If
    
    CountColumns
    
    If InitDownloadInfo() > 0 Then
        Unload Me
        Exit Sub
    End If
    
    'JM 12-18-2015: need to call this here because the grids are getting loaded before showing the form
    FixFormControls Me, ALT_GRID_ROW_COLOR
    
    InitGrid
    LoadGrid
    lblDownload = m.strSizeLabel & " " & m.strSize & " " & m.strKbMb
    
    cmdCancel.Visible = Not m.bHideCancel
    
    If bHiddenSave Then
        SaveDownloadInfo
        m.bHideCancel = False
        Unload Me
    Else
        CenterTheForm Me
        ShowForm Me, eForm_Modal, , , ALT_GRID_ROW_COLOR
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilterDownload.ShowMe", eGDRaiseError_Raise

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    SaveDownloadInfo
    m.bHideCancel = False
    Unload Me

End Sub

Private Sub fgDownload_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If Col = eCfgDailyField Or Col = eCfgIntradayField Then
        With fgDownload
            If .Cell(flexcpChecked, Row, Col) = 1 Or .Cell(flexcpChecked, Row, Col) = 2 Then
                ToggleCheckbox
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilterDownload.fgDownload.AfterEdit", eGDRaiseError_Raise

End Sub

Private Sub fgDownload_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Cancel = True

End Sub

Private Sub fgDownload_Click()
On Error GoTo ErrSection:

    With fgDownload
        If .Col = eCfgDailyField Or .Col = eCfgIntradayField Then
            If .Cell(flexcpChecked, .Row, .Col) = 1 Or .Cell(flexcpChecked, .Row, .Col) = 2 Then
                ToggleCheckbox
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilterDownload.fgDownload.Click", eGDRaiseError_Raise

End Sub

Private Sub Form_Load()

    Me.Icon = Picture16(ToolbarIcon("ID_Download"), , True)
    
    g.Styler.StyleForm Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = m.bHideCancel

End Sub

Private Sub Form_Resize()
On Error GoTo ErrSection:

    If LimitFormSize(Me, fraButtons.Width + 100, 2000) Then Exit Sub
    
    With fgDownload
        If Len(lblGeneralInfo.Caption) > 0 Then
            'fraDwnloadSize.Move Me.ScaleWidth - fraDwnloadSize.Width, lblGeneralInfo.Height
            .Move 0, lblGeneralInfo.Height, Me.ScaleWidth, Me.ScaleHeight - (fraButtons.Height + lblGeneralInfo.Height + 50)
        Else
            'fraDwnloadSize.Move Me.ScaleWidth - fraDwnloadSize.Width, 0
            .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - (fraButtons.Height + 50)
        End If
        If m.nMaxCol = eCfgToolTipField + 1 Then
            .ColWidth(eCfgDailyField) = 1000
            .ColWidth(eCfgIntradayField) = 1000
            .ColWidth(eCfgDescField) = .ClientWidth - 2010
        Else
            .AutoSize 0, m.nMaxCol
        End If
    End With
    'center the buttons
    With fraButtons
        .Top = Me.ScaleHeight - .Height
        If cmdCancel.Visible Then
            .Left = Me.ScaleWidth / 2 - .Width / 2
        Else
            .Left = Me.ScaleWidth / 2 - .Width / 2 + cmdOK.Width / 2
        End If
    End With
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilterDownload.Form.Resize", eGDRaiseError_Raise

End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:
    
    With fgDownload
        .Redraw = flexRDNone
        SetupGrid Me.fgDownload, eGridMode_Grid
        .ExplorerBar = flexExNone
        .ScrollBars = flexScrollBarVertical
        .Editable = flexEDKbdMouse
        .FixedCols = 0
        .FixedRows = 0
        .Rows = 0
        .Cols = m.nMaxCol + 1
        .ColHidden(eCfgTypeField) = True
        .ColHidden(eCfgToolTipField) = True
        .ColHidden(eCfgIntradayField) = m.bExtremeChartsMode
        .ColHidden(m.nMaxCol) = True
        .ColAlignment(eCfgDailyField) = flexAlignCenterCenter
        .ColAlignment(eCfgIntradayField) = flexAlignCenterCenter
        If m.nMaxCol = eCfgToolTipField + 1 Then
            .ExtendLastCol = False
        Else
            .ExtendLastCol = True
            .AutoSizeMode = flexAutoSizeColWidth
        End If
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilterDownload.InitGrid", eGDRaiseError_Raise

End Sub

Private Sub CountColumns()
On Error GoTo ErrSection:

    Dim i&, j&, k&
    Dim nFields&
    Dim strText$
    Dim aText As New cGdArray
    
    j = m.aCfgFile.Size - 1
    'remove all comments and blank lines from array
    For i = j To 0 Step -1
        strText = m.aCfgFile(i)
        If Left(strText, 1) = "'" Then
            m.aCfgFile.Remove i
        Else
            aText.SplitFields strText, vbTab
            If aText.Size = 0 Then
                m.aCfgFile.Remove i
            End If
        End If
    Next
    
    j = m.aCfgFile.Size
    For i = 0 To j
        strText = m.aCfgFile(i)
        aText.SplitFields strText, vbTab
        If aText.Size > k Then k = aText.Size
    Next
    
    m.nMaxCol = k

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilterDownload.CountColumns", eGDRaiseError_Raise

End Sub

Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim i&, j&
    Dim strLower$
    Dim aText As New cGdArray
    
    j = m.aCfgFile.Size - 1
    m.strSupportMsg = ""
    
    fgDownload.Redraw = flexRDNone
    
    For i = 0 To j
        aText.SplitFields m.aCfgFile(i), vbTab
        If aText.Size > 0 Then
            strLower = LCase(Trim(aText(0)))    'splitfields does not trim spaces
            If strLower = "title" Then
                Me.Caption = aText(1)
            ElseIf strLower = "info" Then
                lblGeneralInfo.Caption = Trim(aText(1))
            ElseIf strLower = "header" Then
                LoadHeader aText
            ElseIf strLower = "downloadsize" Then
                m.strSizeLabel = Trim(aText(1))
            ElseIf strLower = "supportmsg" Then
                m.strSupportMsg = Trim(aText(1))
            ElseIf strLower = "item" Then
                LoadItem aText, i
            End If
        End If
    Next
    
    fgDownload.Redraw = flexRDBuffered
    
    If Len(m.strSupportMsg) = 0 Then
        m.strSupportMsg = kSupportMsg
    End If
    m.strSupportMsg = Replace(m.strSupportMsg, "Genesis", GetProvidedProperty("CompanyName", , True))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilterDownload.LoadGrid", eGDRaiseError_Raise

End Sub

Private Sub LoadHeader(aFields As cGdArray)
On Error GoTo ErrSection:

    Dim i&

    With fgDownload
        .Rows = .Rows + 1
        For i = 0 To aFields.Size - 1
            .TextMatrix(.Rows - 1, i) = aFields(i)
        Next
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, m.nMaxCol) = ALT_GRID_ROW_COLOR
        'last col holds index into config data array which is not needed for header
        .TextMatrix(.Rows - 1, m.nMaxCol) = -1
        
        If m.bExtremeChartsMode And .Rows > 10 Then
            .RowHidden(.Rows - 1) = True
        Else
            .RowHidden(.Rows - 1) = False
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilterDownload.LoadHeader", eGDRaiseError_Raise

End Sub

Private Sub LoadItem(aFields As cGdArray, ByVal CfgArrayIdx&)
On Error GoTo ErrSection:

    Dim i&
    Dim aDaily As New cGdArray
    Dim aIntraday As New cGdArray
    Dim nCheckboxValue&

    If aFields.Size > eCfgDescField Then
        With fgDownload
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, eCfgDescField) = aFields(eCfgDescField)
            aDaily.SplitFields aFields(eCfgDailyField), "|"
            aIntraday.SplitFields aFields(eCfgIntradayField), "|"
            If aDaily.Size >= 3 Then
                nCheckboxValue = CheckBoxValue(aDaily(1))
                .Cell(flexcpChecked, .Rows - 1, eCfgDailyField) = nCheckboxValue
                .Cell(flexcpPictureAlignment, .Rows - 1, eCfgDailyField) = flexPicAlignCenterCenter
            End If
            If aIntraday.Size >= 3 Then
                nCheckboxValue = CheckBoxValue(aIntraday(1))
                .Cell(flexcpChecked, .Rows - 1, eCfgIntradayField) = nCheckboxValue
                .Cell(flexcpPictureAlignment, .Rows - 1, eCfgIntradayField) = flexPicAlignCenterCenter
            End If
            'save config data array index into grid
            If aDaily.Size >= 3 Or aIntraday.Size >= 3 Then
                .TextMatrix(.Rows - 1, m.nMaxCol) = CfgArrayIdx
            Else
                .TextMatrix(.Rows - 1, m.nMaxCol) = -1
            End If
            For i = eCfgToolTipField + 1 To aFields.Size - 1
                .TextMatrix(.Rows - 1, i) = aFields(i)
            Next
            
            i = False
            If m.bExtremeChartsMode Then
                If aDaily(5) <> "B" Then
                    i = True
                End If
            End If
            .RowHidden(.Rows - 1) = i
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilterDownload.LoadItem", eGDRaiseError_Raise

End Sub

Private Sub ToggleCheckbox()
On Error GoTo ErrSection:

    Dim i&
    Dim nDefault&, nHasModule&
    Dim aFields As New cGdArray
    Dim aDaily As New cGdArray
    Dim aIntraday As New cGdArray
    
    With fgDownload
        i = Val(.TextMatrix(.Row, m.nMaxCol))   'get index into config data array
        If i < 0 Or (.Col <> eCfgDailyField And .Col <> eCfgIntradayField) Then
            Exit Sub
        End If
    End With
    
    With fgDownload
        
        aFields.SplitFields m.aCfgFile(i), vbTab
        aDaily.SplitFields aFields(eCfgDailyField), "|"
        aIntraday.SplitFields aFields(eCfgIntradayField), "|"
        
        If .Col = eCfgDailyField Then
            nDefault = TableData(aDaily(1), eTableDownloadDefault)
            nHasModule = TableData(aDaily(1), eTableHasModule)
            If nDefault >= 0 And nHasModule = 1 Then ToggleDaily aDaily, aIntraday
        ElseIf .Col = eCfgIntradayField Then
            nDefault = TableData(aIntraday(1), eTableDownloadDefault)
            nHasModule = TableData(aIntraday(1), eTableHasModule)
            If nDefault >= 0 And nHasModule = 1 Then ToggleIntraday aDaily, aIntraday
        End If
        
        If nDefault = -1 Or nHasModule = 0 Then
            InfBox m.strSupportMsg, "I", , Me.Caption
        Else
            CalcDownloadSize
        End If
            
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilterDownload.ToggleCheckbox", eGDRaiseError_Raise

End Sub

Private Sub ToggleDaily(aDaily As cGdArray, aIntraday As cGdArray)
On Error GoTo ErrSection:

    With fgDownload
        If .Cell(flexcpChecked, .Row, eCfgDailyField) = 1 Then
            .Cell(flexcpChecked, .Row, eCfgDailyField) = 2
            TableData(aDaily(1), eTableExclude) = 1
            TableData(aDaily(1), eTableInclude) = 0
            If .Cell(flexcpChecked, .Row, eCfgIntradayField) = 1 Then
                ToggleIntraday aDaily, aIntraday
            End If
        ElseIf .Cell(flexcpChecked, .Row, eCfgDailyField) = 2 Then
            .Cell(flexcpChecked, .Row, eCfgDailyField) = 1
            TableData(aDaily(1), eTableExclude) = 0
            TableData(aDaily(1), eTableInclude) = 1
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilterDownload.ToggleDaily", eGDRaiseError_Raise

End Sub

Private Sub ToggleIntraday(aDaily As cGdArray, aIntraday As cGdArray)
On Error GoTo ErrSection:

    With fgDownload
        If .Cell(flexcpChecked, .Row, eCfgIntradayField) = 1 Then
            .Cell(flexcpChecked, .Row, eCfgIntradayField) = 2
            TableData(aIntraday(1), eTableExclude) = 1
            TableData(aIntraday(1), eTableInclude) = 0
        ElseIf .Cell(flexcpChecked, .Row, eCfgIntradayField) = 2 Then
            .Cell(flexcpChecked, .Row, eCfgIntradayField) = 1
            TableData(aIntraday(1), eTableExclude) = 0
            TableData(aIntraday(1), eTableInclude) = 1
            If .Cell(flexcpChecked, .Row, eCfgDailyField) <> 1 Then
                ToggleDaily aDaily, aIntraday
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilterDownload.ToggleIntraday", eGDRaiseError_Raise

End Sub

Private Function InitDownloadInfo() As Long
'returns: 0=success
'         1=table has no records
'         2=config file has invalid default download
'         3=config file has duplicates
'         4=INI file has conflicting include/exclude information
On Error GoTo ErrSection:

    Dim i&, j&
    Dim strText$
    Dim aFields As New cGdArray
    Dim aDaily As New cGdArray
    Dim aIntraday As New cGdArray
    Dim aErrMsgs As New cGdArray
    Dim bError As Boolean
    
    aErrMsgs.Add ""
    aErrMsgs.Add "The CONFIG file has no download data."
    aErrMsgs.Add "The CONFIG file has invalid default download data: "
    aErrMsgs.Add "The CONFIG file has duplicate download data: "
    aErrMsgs.Add "The INI file has conflicting include/exclude data."
    
    Set m.tbDownloadInfo = Nothing
    Set m.tbDownloadInfo = New cGdTable
    'create fields for download info table (first 4 fields are X/Y/Z/S from config file)
    m.tbDownloadInfo.CreateField eGDARRAY_Strings, 0, "AuthorizationString"
    m.tbDownloadInfo.CreateField eGDARRAY_Strings, 1, "DownloadString"
    m.tbDownloadInfo.CreateField eGDARRAY_Longs, 2, "DownloadDefault"
    m.tbDownloadInfo.CreateField eGDARRAY_Doubles, 3, "DownloadSize"
    m.tbDownloadInfo.CreateField eGDARRAY_Longs, 4, "ExcludeOverride"
    m.tbDownloadInfo.CreateField eGDARRAY_Longs, 5, "IncludeOverride"
    m.tbDownloadInfo.CreateField eGDARRAY_Longs, 6, "HasModule"
    m.tbDownloadInfo.CreateField eGDARRAY_Strings, 7, "IgnoreSymbols"
    
    'populate table with information from config file
    For i = 0 To m.aCfgFile.Size - 1
        aFields.SplitFields m.aCfgFile(i), vbTab
        If aFields(eCfgTypeField) = "item" Then
            aDaily.SplitFields aFields(eCfgDailyField), "|"
            aIntraday.SplitFields aFields(eCfgIntradayField), "|"
            If aDaily.Size >= 3 Then
                m.tbDownloadInfo.AddRecord ""
                j = m.tbDownloadInfo.NumRecords - 1
                m.tbDownloadInfo(0, j) = Trim(aDaily(0))
                m.tbDownloadInfo(1, j) = Trim(aDaily(1))
                aDaily(2) = Trim(aDaily(2))
                If aDaily(2) = "0" Or aDaily(2) = "1" Or aDaily(2) = "-1" Then
                    m.tbDownloadInfo(2, j) = aDaily(2)
                Else
                    InfBox aErrMsgs(2) & aFields(eCfgDailyField), "Error", , Me.Caption
                    bError = True
                    Exit For
                End If
                If aDaily.Size > 3 And Val(aDaily(3)) > 0 Then
                    m.tbDownloadInfo(3, j) = Val(aDaily(3))
                Else
                    m.tbDownloadInfo(3, j) = 0
                End If
                If aDaily.Size > 4 And Len(aDaily(4)) > 0 Then
                    m.tbDownloadInfo(7, j) = Trim(aDaily(4))
                Else
                    m.tbDownloadInfo(7, j) = ""
                End If
                m.tbDownloadInfo(4, j) = 0
                m.tbDownloadInfo(5, j) = 0
                m.tbDownloadInfo(6, j) = 0
            End If
            If aIntraday.Size >= 3 Then
                m.tbDownloadInfo.AddRecord ""
                j = m.tbDownloadInfo.NumRecords - 1
                m.tbDownloadInfo(0, j) = Trim(aIntraday(0))
                m.tbDownloadInfo(1, j) = Trim(aIntraday(1))
                aIntraday(2) = Trim(aIntraday(2))
                If aIntraday(2) = "0" Or aIntraday(2) = "1" Or aIntraday(2) = "-1" Then
                    m.tbDownloadInfo(2, j) = aIntraday(2)
                Else
                    InfBox aErrMsgs(2) & aFields(eCfgIntradayField), "Error", , Me.Caption
                    bError = True
                    Exit For
                End If
                If aIntraday.Size > 3 And Val(aIntraday(3)) > 0 Then
                    m.tbDownloadInfo(3, j) = Val(aIntraday(3))
                Else
                    m.tbDownloadInfo(3, j) = 0
                End If
                m.tbDownloadInfo(4, j) = 0
                m.tbDownloadInfo(5, j) = 0
                m.tbDownloadInfo(6, j) = 0
            End If
        End If
    Next
    
    'check that something was read into the table
    If m.tbDownloadInfo.NumRecords = 0 Then
        InfBox aErrMsgs(1), "Error", , Me.Caption
        InitDownloadInfo = 1
        Exit Function
    ElseIf bError Then
        InitDownloadInfo = 2        'invalid download default
        Exit Function
    End If
    
    'check for duplicate secondary/download code
    Set m.aIdxByDwnloadStr = m.tbDownloadInfo.CreateSortedIndex(eTableDownloadStr)
    For i = 1 To m.aIdxByDwnloadStr.Size - 1
        If m.tbDownloadInfo(eTableDownloadStr, m.aIdxByDwnloadStr(i)) = m.tbDownloadInfo(eTableDownloadStr, m.aIdxByDwnloadStr(i - 1)) Then
            InfBox aErrMsgs(3) & m.tbDownloadInfo(eTableDownloadStr, m.aIdxByDwnloadStr(i)), "Error", , Me.Caption
            bError = True
            Exit For
        End If
    Next
    
    If bError Then
        InitDownloadInfo = 3    'config file has duplicates
        Exit Function
    End If
    
    
    m.bHideCancel = True '(default, unless one of the 2 strings are in the INI file)
    
    'set exclude field with exclusion info from INI file
    strText = GetIniFileProperty("DownloadExclude", "", "General", g.strIniFile)
    If Len(strText) > 0 Then
        m.bHideCancel = False
        aFields.SplitFields strText
        aFields.Sort eGdSort_DeleteDuplicates, 0, -1
        For i = 0 To aFields.Size - 1
            m.tbDownloadInfo.SearchAsIndex m.aIdxByDwnloadStr, eTableDownloadStr, aFields(i), j
            If aFields(i) = m.tbDownloadInfo(eTableDownloadStr, m.aIdxByDwnloadStr(j)) Then
                m.tbDownloadInfo(eTableExclude, m.aIdxByDwnloadStr(j)) = 1
            End If
        Next
    End If
    
    'set include field with inclusion info from INI file
    strText = GetIniFileProperty("DownloadInclude", "", "General", g.strIniFile)
    If Len(strText) > 0 Then
        m.bHideCancel = False
        aFields.SplitFields strText
        aFields.Sort eGdSort_DeleteDuplicates, 0, -1
        For i = 0 To aFields.Size - 1
            m.tbDownloadInfo.SearchAsIndex m.aIdxByDwnloadStr, eTableDownloadStr, aFields(i), j
            If aFields(i) = m.tbDownloadInfo(eTableDownloadStr, m.aIdxByDwnloadStr(j)) Then
                If m.tbDownloadInfo(eTableExclude, m.aIdxByDwnloadStr(j)) = 0 Then
                    m.tbDownloadInfo(eTableInclude, m.aIdxByDwnloadStr(j)) = 1
                Else
                    bError = True
                    Exit For
                End If
            End If
        Next
    Else
        m.bHideCancel = True
    End If

    If bError Then
        InfBox aErrMsgs(4), "Error", , Me.Caption
        InitDownloadInfo = 4        'conflicting exclude/include info in INI file
        Exit Function
    End If
    
    'set has module field
    For i = 0 To m.tbDownloadInfo.NumRecords - 1
        m.tbDownloadInfo(eTableHasModule, i) = Abs(HasModule(m.tbDownloadInfo(eTableAuthStr, i)))
    Next
    
    'download size
    CalcDownloadSize
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmFilterDownload.InitDownloadInfo", eGDRaiseError_Raise

End Function

Private Property Get CheckBoxValue(ByVal strDownload$, _
    Optional ByVal nIdx& = -1) As Long
On Error GoTo ErrSection:

    Dim i&, j&
    Dim nDefault&, nInclude&, nExclude&, nHasModule&
        
    If nIdx >= 0 Then
        j = nIdx
    Else
        m.tbDownloadInfo.SearchAsIndex m.aIdxByDwnloadStr, eTableDownloadStr, strDownload, i
        j = m.aIdxByDwnloadStr(i)
    End If
    If m.tbDownloadInfo(eTableDownloadStr, j) = strDownload Then
        nDefault = m.tbDownloadInfo(eTableDownloadDefault, j)
        nInclude = m.tbDownloadInfo(eTableInclude, j)
        nExclude = m.tbDownloadInfo(eTableExclude, j)
        nHasModule = m.tbDownloadInfo(eTableHasModule, j)
        
        If nDefault = -1 Then
            If nHasModule Then
                i = 1
            Else
                i = 2
            End If
        ElseIf nHasModule = 0 Then
            i = 2
        ElseIf nInclude = 1 Then
            i = 1
        ElseIf nExclude = 1 Then
            i = 2
        ElseIf nDefault = 0 Then
            i = 2
        ElseIf nDefault = 1 Then
            i = 1
        Else
            i = 0
        End If
    Else
        i = 0       'theoretically should never get here
    End If
    
    CheckBoxValue = i

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmFilterDownload.CheckBoxValue.Get", eGDRaiseError_Raise

End Property

Private Property Get TableData(ByVal strDownload$, ByVal eField As eTableFields) As Long
On Error GoTo ErrSection:

    Dim i&, j&
    
    If eField = eTableAuthStr And eField = eTableDownloadStr Then
        'these are string data type (will add property for these later if needed)
        i = -1
    Else
        m.tbDownloadInfo.SearchAsIndex m.aIdxByDwnloadStr, eTableDownloadStr, strDownload, i
        j = m.aIdxByDwnloadStr(i)
        If m.tbDownloadInfo(eTableDownloadStr, j) = strDownload Then
            i = m.tbDownloadInfo(eField, j)
        Else
            i = -1
        End If
    End If
    
    TableData = i

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmFilterDownload.TableData.Get", eGDRaiseError_Raise

End Property

Private Property Let TableData(ByVal strDownload$, ByVal eField As eTableFields, ByVal nValue&)
On Error GoTo ErrSection:

    Dim i&, j&
    
    If eField = eTableExclude Or eField = eTableInclude Then
        'for now these are the only two fields that should be editable
        m.tbDownloadInfo.SearchAsIndex m.aIdxByDwnloadStr, eTableDownloadStr, strDownload, i
        j = m.aIdxByDwnloadStr(i)
        If m.tbDownloadInfo(eTableDownloadStr, j) = strDownload Then
            m.tbDownloadInfo(eField, j) = nValue
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmFilterDownload.TableData.Let", eGDRaiseError_Raise

End Property

Private Sub SaveDownloadInfo()
On Error GoTo ErrSection:

    Dim i&, bIgnoreSym As Boolean
    Dim strInclude$, strExclude$, strIgnoreSym$
    Dim nDefault&, nHasModule&, nInclude&, nExclude
    
    For i = 0 To m.tbDownloadInfo.NumRecords - 1
        nDefault = m.tbDownloadInfo(eTableDownloadDefault, i)
        nHasModule = m.tbDownloadInfo(eTableHasModule, i)
        nInclude = m.tbDownloadInfo(eTableInclude, i)
        nExclude = m.tbDownloadInfo(eTableExclude, i)
        bIgnoreSym = False
        If nHasModule = 0 Then
            bIgnoreSym = True
        ElseIf nDefault <> -1 Then
            If nInclude = 1 Then
                strInclude = strInclude & m.tbDownloadInfo(eTableDownloadStr, i) & ","
            ElseIf nExclude = 1 Then
                strExclude = strExclude & m.tbDownloadInfo(eTableDownloadStr, i) & ","
                bIgnoreSym = True
            ElseIf nInclude = 0 And nExclude = 0 Then
                If nDefault = 1 Then
                    strInclude = strInclude & m.tbDownloadInfo(eTableDownloadStr, i) & ","
                ElseIf nDefault = 0 Then
                    strExclude = strExclude & m.tbDownloadInfo(eTableDownloadStr, i) & ","
                    bIgnoreSym = True
                End If
            End If
        End If
        If bIgnoreSym And Len(m.tbDownloadInfo(eTableIgnoreSym, i)) > 0 Then
            strIgnoreSym = strIgnoreSym & m.tbDownloadInfo(eTableIgnoreSym, i) & ","
        End If
    Next
        
    strInclude = Trim(strInclude)
    strExclude = Trim(strExclude)
    If Len(strInclude) > 0 Then
        strInclude = Left(strInclude, Len(strInclude) - 1) 'remove trailing comma
    Else
        strInclude = ""
    End If
    If Len(strExclude) > 0 Then
        strExclude = Left(strExclude, Len(strExclude) - 1)
    Else
        strExclude = ""
    End If
    If Len(strIgnoreSym) > 0 Then
        strIgnoreSym = Left(strIgnoreSym, Len(strIgnoreSym) - 1)
    Else
        strIgnoreSym = ""
    End If
    
    SetIniFileProperty "DownloadInclude", strInclude, "General", g.strIniFile
    SetIniFileProperty "DownloadExclude", strExclude, "General", g.strIniFile
    SetIniFileProperty "IgnoreSymbols", strIgnoreSym, "General", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilterDownload.SaveDownloadInfo", eGDRaiseError_Raise

End Sub

Private Sub CalcDownloadSize()
On Error GoTo ErrSection:

    Dim i&, j&
    Dim dSize#, dSizeTotal#
    Dim strSize$, strKbMb$

    dSizeTotal = 0#
    For i = 0 To m.tbDownloadInfo.NumRecords - 1
        If m.tbDownloadInfo(eTableHasModule, i) <> 0 Then
            dSize = m.tbDownloadInfo(eTableDownloadSize, i)
            If dSize > 0 Then
                j = CheckBoxValue(m.tbDownloadInfo(eTableDownloadStr, i), i)
                If j = 1 Then
                    dSizeTotal = dSizeTotal + dSize
                End If
            End If
        End If
    Next
    
    If dSizeTotal > 1000 Then
        m.strSize = FormatNum(dSizeTotal / 1000, -2)
        m.strKbMb = "MB"
    Else
        m.strSize = FormatNum(dSizeTotal, -2)
        m.strKbMb = "KB"
    End If
    
    lblDownload = m.strSizeLabel & " " & m.strSize & " " & m.strKbMb
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFilterDownload.CalcDownloadSize", eGDRaiseError_Raise

End Sub

