VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmPrint 
   Caption         =   "Report Manager"
   ClientHeight    =   6525
   ClientLeft      =   3135
   ClientTop       =   1500
   ClientWidth     =   7170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin ActiveToolBars.SSActiveToolBars Toolbar1 
      Left            =   1770
      Top             =   315
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   7
      Tools           =   "frmPrint.frx":0000
      ToolBars        =   "frmPrint.frx":3419
   End
   Begin VB.CommandButton Corner 
      Caption         =   "Command1"
      Height          =   525
      Left            =   6360
      TabIndex        =   2
      Top             =   6015
      Visible         =   0   'False
      Width           =   825
   End
   Begin vsOcx6LibCtl.vsElastic vsElastic1 
      Height          =   6525
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   11509
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      Appearance      =   1
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   1
      GridCols        =   1
      _GridInfo       =   $"frmPrint.frx":3544
      Begin VSPrinter7LibCtl.VSPrinter vp 
         Height          =   6345
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   6990
         _cx             =   12330
         _cy             =   11192
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         MousePointer    =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _ConvInfo       =   1
         AutoRTF         =   -1  'True
         Preview         =   -1  'True
         DefaultDevice   =   0   'False
         PhysicalPage    =   -1  'True
         AbortWindow     =   -1  'True
         AbortWindowPos  =   0
         AbortCaption    =   "Printing..."
         AbortTextButton =   "Cancel"
         AbortTextDevice =   "on the %s on %s"
         AbortTextPage   =   "Now printing Page %d of"
         FileName        =   ""
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         MarginHeader    =   0
         MarginFooter    =   0
         IndentLeft      =   0
         IndentRight     =   0
         IndentFirst     =   0
         IndentTab       =   720
         SpaceBefore     =   0
         SpaceAfter      =   0
         LineSpacing     =   100
         Columns         =   1
         ColumnSpacing   =   180
         ShowGuides      =   2
         LargeChangeHorz =   300
         LargeChangeVert =   300
         SmallChangeHorz =   30
         SmallChangeVert =   30
         Track           =   0   'False
         ProportionalBars=   -1  'True
         Zoom            =   34.9431818181818
         ZoomMode        =   3
         ZoomMax         =   400
         ZoomMin         =   10
         ZoomStep        =   25
         EmptyColor      =   -2147483636
         TextColor       =   0
         HdrColor        =   0
         BrushColor      =   0
         BrushStyle      =   0
         PenColor        =   0
         PenStyle        =   0
         PenWidth        =   0
         PageBorder      =   0
         Header          =   ""
         Footer          =   ""
         TableSep        =   "|;"
         TableBorder     =   7
         TablePen        =   0
         TablePenLR      =   0
         TablePenTB      =   0
         NavBar          =   2
         NavBarColor     =   -2147483633
         ExportFormat    =   0
         URL             =   ""
         Navigation      =   3
         NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCaller        As Object

Property Let CallingForm(pData As Object)
    Set mCaller = pData
End Property

Private Sub Form_Activate()
On Error Resume Next
    vp.SetFocus
End Sub

Private Sub Form_Load()
    'Dim w   As String
    'If gAppSettings.frmPrintsettings = "" Then
        ReSizeMDIChildForm Me, Corner
        CenterTheForm Me
    'Else
    '    w = gAppSettings.frmPrintsettings
    '    SetFormPlacement Me, w, "LHTW"
    'End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Dim w   As String
    'w = GetFormPlacement(Me)
    'gAppSettings.frmPrintsettings = w
    'gAppSettings.Save
    If frmPrint.Visible Then
        Cancel = True
    End If
    Me.Hide
End Sub

Private Sub Form_Resize()
    'With vp
    '    On Error Resume Next
    '    .Move .left, .Top, ScaleWidth - 2 * .left, ScaleHeight - .Top - .left
    'End With
End Sub

Private Sub Toolbar1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:
    Dim RetVal  As Variant
    
    Select Case Tool.ID
        Case "ID_Print"
            If vp.PrintDialog(pdPrint) Then
                vp.PrintDoc
            End If
            
        Case "ID_PrinterSetup"
            If vp.PrintDialog(pdPrinterSetup) = True Then
                mCaller.RunReport
            End If
            
        Case "ID_PageSetup"
            If vp.PrintDialog(pdPageSetup) = True Then
                mCaller.RunReport
            End If
            
        Case "ID_Leave": Unload Me
    End Select
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "Print.Toolbar1.ToolClick", eGDRaiseError_Show, frmReports.AppPath
    Resume ErrExit:

End Sub

