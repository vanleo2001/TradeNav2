VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmDelete 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraShowMore 
      Height          =   3675
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3315
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
      Caption         =   "frmDelete.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmDelete.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDelete.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniFrameWL fraButtons 
         Height          =   1935
         Left            =   2100
         TabIndex        =   5
         Top             =   300
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
         Caption         =   "frmDelete.frx":0068
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmDelete.frx":0094
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDelete.frx":00B4
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
            Default         =   -1  'True
            Height          =   375
            Index           =   0
            Left            =   60
            TabIndex        =   1
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
            Caption         =   "frmDelete.frx":00D0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmDelete.frx":00FE
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmDelete.frx":011E
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
            Height          =   375
            Index           =   0
            Left            =   60
            TabIndex        =   2
            Top             =   480
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
            Caption         =   "frmDelete.frx":013A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmDelete.frx":0168
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmDelete.frx":0188
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniLabelXP lblNote 
            Height          =   855
            Left            =   60
            Top             =   1020
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
            Caption         =   "frmDelete.frx":01A4
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmDelete.frx":0238
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmDelete.frx":0258
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid fgList 
         Height          =   3435
         Left            =   0
         TabIndex        =   0
         Top             =   240
         Width           =   2055
         _cx             =   3625
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
      Begin HexUniControls.ctlUniLabelXP lblDesc 
         Height          =   255
         Left            =   60
         Top             =   0
         Width           =   1875
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
         Caption         =   "frmDelete.frx":0274
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDelete.frx":02C8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDelete.frx":02E8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraConfirm 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   4260
      Width           =   3315
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
      Caption         =   "frmDelete.frx":0304
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmDelete.frx":0324
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDelete.frx":0344
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkShowMore 
         Height          =   220
         Left            =   0
         TabIndex        =   9
         Top             =   840
         Width           =   2715
         _ExtentX        =   4789
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
         Caption         =   "frmDelete.frx":0360
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmDelete.frx":03BC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmDelete.frx":03DC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniFrameWL fraButtons2 
         Height          =   375
         Left            =   540
         TabIndex        =   6
         Top             =   1200
         Width           =   2235
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
         Caption         =   "frmDelete.frx":03F8
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmDelete.frx":0424
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDelete.frx":0444
         RightToLeft     =   0   'False
         Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   8
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
            Caption         =   "frmDelete.frx":0460
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmDelete.frx":048E
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmDelete.frx":04AE
            RightToLeft     =   0   'False
         End
         Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
            Cancel          =   -1  'True
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   7
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
            Caption         =   "frmDelete.frx":04CA
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            ShowFocus       =   -1  'True
            Tristate        =   0   'False
            Pressed         =   0   'False
            Tip             =   "frmDelete.frx":04F8
            Style           =   -1
            RoundedBorders  =   -1  'True
            xTranspColor    =   0
            yTranspColor    =   0
            MousePointer    =   0
            MouseIcon       =   "frmDelete.frx":0518
            RightToLeft     =   0   'False
         End
      End
      Begin VB.PictureBox ico 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   550
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   4
         Top             =   0
         Width           =   550
      End
      Begin HexUniControls.ctlUniLabelXP lblMessage 
         Height          =   555
         Left            =   600
         Top             =   0
         Width           =   2655
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
         Caption         =   "frmDelete.frx":0534
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmDelete.frx":0554
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmDelete.frx":0574
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmDelete.frm
'' Description: Allow the user to delete one or more items from a list
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eDeleteFormMode
    eDeleteFormMode_Confirm = 0
    eDeleteFormMode_ShowMore
End Enum

Private Type mPrivate
    bOK As Boolean
    nMode As eDeleteFormMode
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Array of items to show in the list, Item to select by default
'' Returns:     True if OK, False if Cancel
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal astrList As cGdArray, Optional ByVal strSelect As String = "") As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of values to return
    Dim lIndex As Long                  ' Index into a for loop
    Dim hIcon As Long                   ' Handle to the loaded icon
    Dim lRC As Long                     ' Return code from functions
    
    If astrList.Size > 0 Then
        fgList.Redraw = flexRDNone
        InitGrid
        LoadGrid astrList, strSelect
        fgList.Redraw = flexRDBuffered
        
        'RH commented out ico.BackColor = g.nColorTheme       'BackColor
        hIcon = LoadIconNum(0, 32514)
        If hIcon Then
            lRC = DrawIcon(ico.hDC, 0, 0, hIcon)
            lRC = DestroyIcon(hIcon)
        End If
    
        If Len(Trim(strSelect)) = 0 Then
            ChangeMode eDeleteFormMode_ShowMore
        Else
            lblMessage.Caption = "Are you sure you want to delete" & vbCrLf & strSelect & "?"
            ChangeMode eDeleteFormMode_Confirm
        End If
        
        ShowForm Me, True
        
        If m.bOK Then
            Set astrReturn = New cGdArray
            astrReturn.Create eGDARRAY_Strings
            
            For lIndex = 0 To fgList.SelectedRows - 1
                astrReturn.Add fgList.TextMatrix(fgList.SelectedRow(lIndex), 1)
            Next lIndex
            
            Set ShowMe = astrReturn
        Else
            Set ShowMe = Nothing
        End If
    End If

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmDelete.ShowMe", eGDRaiseError_Raise
    
End Function

Private Sub chkShowMore_Click()
On Error GoTo ErrSection:

    If chkShowMore.Value = vbChecked Then
        ChangeMode eDeleteFormMode_ShowMore
    Else
        ChangeMode eDeleteFormMode_Confirm
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDelete.chkShowMore.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow the user to cancel the deletion
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click(Index As Integer)
On Error GoTo ErrSection:

    m.bOK = False
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDelete.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDelete_Click
'' Description: Allow the user to continue with the deletion
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDelete_Click(Index As Integer)
On Error GoTo ErrSection:

    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDelete.cmdDelete.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Dim lIndex As Long

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    ElseIf Shift = vbCtrlMask And UCase(Chr(KeyCode)) = "A" Then
        ' Ctrl-A: hot-key to select all
        With fgList
            If .Rows < 5000 And .AllowSelection Then
                For lIndex = .FixedRows To .Rows - 1
                    .IsSelected(lIndex) = True
                Next
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDelete.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize some stuff when the form is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Caption = "Confirmation"
    Icon = Picture16("kBlank")
    
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDelete.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user hits the X, act as if they cancelled the delete
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
    RaiseError "frmDelete.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move/Resize controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lMinWidth As Long               ' Minimum scale width to allow for the form
    Dim lMinHeight As Long              ' Minimum scale height to allow for the form
    
    Select Case m.nMode
        Case eDeleteFormMode_Confirm
            'lMinWidth = fraConfirm.Width + fraConfirm.Left * 2
            'lMinHeight = fraConfirm.Height + fraConfirm.Height * 2
            'If LimitFormSize(Me, lMinWidth, lMinHeight) Then Exit Sub
            
            'With fraConfirm
            '    .Move .Left, .Top
            'End With
        
        Case eDeleteFormMode_ShowMore
            'lMinWidth = lblDesc.Width + (lblDesc.Left * 2)
            'lMinHeight = fraButtons.Height + fraButtons.Top
            'If LimitFormSize(Me, lMinWidth, lMinHeight) Then Exit Sub
            
            'With fraButtons
            '    .Left = ScaleWidth - .Width
            'End With
            
            'With fgList
            '    .Move .Left, .Top, fraButtons.Left - .Left, ScaleHeight - .Top - .Left
            'End With
    
    End Select

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

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With fgList
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        SetupGrid fgList, eGridMode_List
        .AllowSelection = True
        .Cols = 2
        .FixedCols = 0
        .ColHidden(1) = True
        .Rows = 0
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDelete.InitGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid from the given list
'' Inputs:      List to load the grid from
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid(ByVal astrList As cGdArray, Optional ByVal strSelect As String = "")
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRowSel As Long                 ' Row selected in the grid
    
    With fgList
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        lRowSel = -1&
        For lIndex = 0 To astrList.Size - 1
            .AddItem astrList(lIndex)
            If Trim(Parse(astrList(lIndex), vbTab, 1)) = Trim(strSelect) Then
                .Row = .Rows - 1
                .RowSel = .Rows - 1
                lRowSel = .RowSel
            End If
        Next lIndex
        
        If lRowSel <> -1& Then
            .ShowCell lRowSel, 0
        End If
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDelete.LoadGrid", eGDRaiseError_Raise
    
End Sub

Private Sub ChangeMode(ByVal nMode As eDeleteFormMode)
On Error GoTo ErrSection:

    Dim lWidthDiff As Long
    Dim lHeightDiff As Long
    
    lWidthDiff = Width - ScaleWidth
    lHeightDiff = Height - ScaleHeight
    
    m.nMode = nMode
    
    Select Case nMode
        Case eDeleteFormMode_Confirm
            With fraConfirm
                .Move 120, 120
                .Visible = True
            End With
            fraShowMore.Visible = False
            With Me
                .Move .Left, .Top, fraConfirm.Width + (fraConfirm.Left * 2) + lWidthDiff, _
                        fraConfirm.Height + (fraConfirm.Top * 2) + lHeightDiff
            End With
            cmdDelete(1).Default = True
            
        Case eDeleteFormMode_ShowMore
            With fraShowMore
                .Move 120, 120
                .Visible = True
            End With
            fraConfirm.Visible = False
            With Me
                .Move .Left, .Top, fraShowMore.Width + (fraShowMore.Left * 2) + lWidthDiff, _
                        fraShowMore.Height + (fraShowMore.Top * 2) + lHeightDiff
            End With
            cmdDelete(0).Default = True
    
    End Select
    
    CenterTheForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDelete.ChangeMode", eGDRaiseError_Raise
    
End Sub



