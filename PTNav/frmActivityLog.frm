VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmActivityLog 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMenu 
      Left            =   4080
      Top             =   1980
   End
   Begin VB.Timer tmrRealtime 
      Left            =   4080
      Top             =   2460
   End
   Begin VSFlex7LCtl.VSFlexGrid fgActivityLog 
      Height          =   2895
      Left            =   900
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
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuTradeHistory 
         Caption         =   "Trade History"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuCheckStatus 
         Caption         =   "Check Status"
      End
      Begin VB.Menu mnuViewJournals 
         Caption         =   "View Journals"
      End
   End
End
Attribute VB_Name = "frmActivityLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmActivityLog.frm
'' Description: Form to display the activity log
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/16/2010   DAJ         Fixed default startup position/size
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    ActivityLog As cActivityLog         ' Activity log object
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow as outside caller to print the grid information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection:

    PrintMe = frmPrintPreview.ShowMe("TNV ActivityLog", Me, , , , , , True)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmActivityLog.PrintMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the print preview form for this grid
'' Inputs:      Arguments passed in from PrintMe
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:
    
    m.ActivityLog.GenerateReport vArgs
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActivityLog.GenerateReport"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddEntry
'' Description: Add an entry to the activity log
'' Inputs:      Source, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddEntry(ByVal strSource As String, ByVal strMessage As String)
On Error GoTo ErrSection:

    m.ActivityLog.Add strSource, strMessage
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActivityLog.AddEntry"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form members when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement string from the ini file
    Dim UI As cActivityLogControls      ' Activity log controls
    
    strPlacement = GetIniFileProperty("frmActivityLog", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        Move 1470, 4080, 15720, 3600
    Else
        SetFormPlacement Me, strPlacement
    End If
    
    g.Styler.StyleForm Me
    
    Caption = "Activity Log (right-click on grid to see options)"
    Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)
    
    Set UI = New cActivityLogControls
    With UI
        Set .frm = Me
        Set .fgGrid = fgActivityLog
        Set .tmrRealtime = tmrRealtime
        Set .tmrMenu = tmrMenu
        
        Set .mnuActivityLog = mnuPopUp
        Set .mnuPrint = mnuPrint
        Set .mnuTradeHistory = mnuTradeHistory
        Set .mnuSettings = mnuSettings
        Set .mnuCheckStatus = mnuCheckStatus
        Set .mnuViewJournals = mnuViewJournals
    End With

    Set m.ActivityLog = New cActivityLog
    m.ActivityLog.Init UI
    
    g.ActivityLogs.Add "frmActivityLog", m.ActivityLog
    
    tmrMenu.Interval = 100
    tmrMenu.Enabled = False
    
    mnuPopUp.Visible = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActivityLog.Form_Load"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, re-attach the grid
'' Inputs:      Cancel Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = vbFormControlMenu Then
        If Not g.ConsoleForms Is Nothing Then
            g.ConsoleForms.ShowForm(eGDConsoleForm_ActivityLog) = False
        End If
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActivityLog.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    With fgActivityLog
        .Move 0, 0, ScaleWidth, ScaleHeight
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up the form members when it is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "frmActivityLog", GetFormPlacement(Me), "Placement", g.strIniFile
    g.ActivityLogs.Remove "frmActivityLog"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmActivityLog.Form_Unload"
    
End Sub

