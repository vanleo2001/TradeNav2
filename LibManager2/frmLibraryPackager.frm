VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Begin VB.Form frmLibraryPackager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Library Packager"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrExport 
      Enabled         =   0   'False
      Left            =   6420
      Top             =   4320
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   435
      Left            =   7980
      TabIndex        =   17
      Top             =   4320
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   8940
      Begin VB.CheckBox chkExpDate 
         Caption         =   "I would like the Exported Library File to &Expire on:"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   3630
         Width           =   3855
      End
      Begin VB.TextBox txtCustID 
         Height          =   315
         Left            =   6540
         TabIndex        =   11
         Top             =   3270
         Width           =   1875
      End
      Begin VB.CheckBox chkCustID 
         Caption         =   "I only want the Genesis User with the following Customer &ID to Import this Library File:"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   3300
         Width           =   6375
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2100
         Width           =   4620
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8160
         TabIndex        =   8
         Top             =   2100
         Width           =   360
      End
      Begin VB.TextBox txtLibraryName 
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1695
         Width           =   4620
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   120
         Picture         =   "frmLibraryPackager.frx":0000
         ScaleHeight     =   2625
         ScaleWidth      =   1935
         TabIndex        =   1
         Top             =   180
         Width           =   1935
      End
      Begin gdOCX.gdSelectDate gdExpDate 
         Height          =   315
         Left            =   4080
         TabIndex        =   13
         Top             =   3600
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
      End
      Begin VB.Label Label3 
         Caption         =   "Library File Security"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   2940
         Width           =   2475
      End
      Begin VB.Label Label5 
         Caption         =   "Export &Path:"
         Height          =   195
         Index           =   1
         Left            =   2190
         TabIndex        =   6
         Top             =   2145
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "&Library to Export:"
         Height          =   195
         Index           =   0
         Left            =   2175
         TabIndex        =   4
         Top             =   1740
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   $"frmLibraryPackager.frx":7F7E
         Height          =   855
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   6645
      End
      Begin VB.Label Label2 
         Caption         =   "Export a Library"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   300
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "E&xport"
      Height          =   435
      Left            =   6900
      TabIndex        =   16
      Top             =   4320
      Width           =   1035
   End
   Begin VSFlex7LCtl.VSFlexGrid vsStatus 
      Height          =   315
      Left            =   135
      TabIndex        =   14
      Top             =   4365
      Visible         =   0   'False
      Width           =   1365
      _cx             =   2408
      _cy             =   556
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
      BackColor       =   0
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   8421504
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483636
      BackColorAlternate=   0
      GridColor       =   -2147483639
      GridColorFixed  =   -2147483639
      TreeColor       =   -2147483632
      FloodColor      =   16711680
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VB.Label lblStatusMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1635
      TabIndex        =   15
      Top             =   4365
      Visible         =   0   'False
      Width           =   5145
   End
End
Attribute VB_Name = "frmLibraryPackager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmLibraryPackager.frm
'' Description: Allow the user to export a library
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/03/2013   DAJ         Move Strategy Baskets into database
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

Private Type mPrivate
    Library As cLibrary                 ' Library object
    bOK As Boolean
    bAutoMode As Boolean
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkCustID_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkCustID_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryPackager.chkCustID.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkExpDate_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkExpDate_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryPackager.chkExpDate.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Unload the form without saving
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:
    
    m.bOK = False
    Me.Hide
    
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    RaiseError "frmLibraryPackager.cmdCancel.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdExport_Click
'' Description: Export the library and unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdExport_Click()
On Error GoTo ErrSection:
    
    Dim astrDep As New cGdArray         ' Array of dependant libraries
    Dim strMessage As String            ' Message to display to the user
    Dim lIndex As Long                  ' Index into a for loop
    Dim strReturn As String             ' Return from the message box
    Dim bUserLibrary As Boolean         ' Is the user library one of them?
    
    strReturn = "E"
    astrDep.Create eGDARRAY_Strings
    
    If EnsureNoUserDeps = True Then
        LibraryDependencies astrDep, m.Library.LibraryID, eGDDependencyFilter_NonBuiltInNonUser
        For lIndex = astrDep.Size - 1 To 0 Step -1
            If Parse(astrDep(lIndex), vbTab, 4) = Str(m.Library.LibraryID) Then
                astrDep.Remove lIndex
            End If
        Next lIndex
        
        astrDep.Sort eGdSort_DeleteDuplicates
        
        If astrDep.Size > 10 Then
            strMessage = "Warning: Items in this library require|functions from other libraries."
            strMessage = strMessage & "||Successfully importing this library will require|the other libraries to exist."
            strReturn = InfBox(strMessage, , "+Export|-Cancel", "Dependant Libraries")
        ElseIf astrDep.Size > 0 Then
            strMessage = "Warning: Items in this library require the following function(s) in other libraries:|"
            For lIndex = 0 To astrDep.Size - 1
                strMessage = strMessage & "|" & Parse(astrDep(lIndex), vbTab, 2) & " (" & Parse(astrDep(lIndex), vbTab, 5) & ")"
            Next lIndex
            strMessage = strMessage & "||Successfully importing this library will require the other libraries to exist."
            
            strReturn = InfBox(strMessage, , "+Export|-Cancel", "Dependant Libraries")
        End If
    Else
        strReturn = "C"
    End If
    
    If strReturn = "E" Then Export
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibraryPackager.cmdExport.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdBrowse_Click
'' Description: Allow the user to browse for a filename
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdBrowse_Click()
On Error GoTo ErrSection:

    txtPath.Text = frmDirList.ShowMe
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibraryPackager.cmdBrowse.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Show the help screen if the user presses F1
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryPackager.Form.KeyDown", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form and controls
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim lExpDays As Long                ' Default number of expiration days

    Icon = Picture16("kLibrary")
    CenterTheForm Me
    
    If Not DirExist(AddSlash(g.strAppPath) & "Libraries") Then MkDir AddSlash(g.strAppPath) & "Libraries"
    txtPath.Text = AddSlash(g.strAppPath) & "Libraries"
    
    txtCustID.Text = GetIniFileProperty("CustIDs", "", "Library", g.strIniFile)
    lExpDays = GetIniFileProperty("ExpDays", 30, "Library", g.strIniFile)
    gdExpDate.Value = Date + lExpDays
    gdExpDate.MinDate = Date
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibraryPackager.Form.Load", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Unload the form if the user hits the 'X'
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
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
    RaiseError "frmLibraryPackager.Form.QueryUnload", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Cleanup on exit
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    tmrExport.Enabled = False
    Set m.Library = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryPackager.Form.Unload", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      ID of the Library to load
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal lLibraryID As Long, Optional ByVal bAuto As Boolean = False) As Boolean
On Error GoTo ErrSection:
    
    m.bAutoMode = bAuto
    
    Set m.Library = New cLibrary
    With m.Library
        .LibraryID = lLibraryID
        .Load
        txtLibraryName.Text = .LibraryName
    End With
    EnableControls
    
    tmrExport.Interval = 1000
    tmrExport.Enabled = bAuto
    ShowForm Me, eForm_Modal, g.frmOwner
    
    ShowMe = m.bOK
    
ErrExit:
    Unload Me
    Exit Function

ErrSection:
    Unload Me
    RaiseError "frmLibraryPackager.ShowMe", eGDRaiseError_Raise, g.strAppPath

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable Controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Enable txtCustID, (chkCustID = vbChecked)
    Enable gdExpDate, (chkExpDate = vbChecked)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryPackager.EnableControls", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Export
'' Description: Export the library
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Export()
On Error GoTo ErrSection:

    Screen.MousePointer = vbHourglass
    
    If Len(txtLibraryName.Text) = 0 Then
        Err.Raise vbObjectError + 1000, , "Please type a library name"
    End If
    
    If Len(txtPath.Text) = 0 Then
        Err.Raise vbObjectError + 1000, , "Please select a file name (and path) to store your library"
    End If
    
    If Len(txtCustID.Text) = 0 And chkCustID = vbChecked Then
        Err.Raise vbObjectError + 1000, , "To secure the Library, please Enter in a Customer ID"
    End If
    
    With m.Library
        .StatusBar = vsStatus
        .Path = txtPath.Text & "\" & txtLibraryName.Text & ".txt"
        .StatusMsg = lblStatusMsg
        
        If chkCustID = vbChecked Then
            .CustomerID = txtCustID.Text
            SetIniFileProperty "CustIDs", txtCustID.Text, "Library", g.strIniFile
        Else
            .CustomerID = ""
        End If
        If chkExpDate = vbChecked Then
            .ExpirationDate = gdExpDate.Value
            SetIniFileProperty "ExpDays", CLng(gdExpDate.Value) - CLng(Date), "Library", g.strIniFile
        Else
            .ExpirationDate = 0
        End If
        
        .Package
    End With
    
    If m.bAutoMode = False Then
        InfBox "The library file:|" & UCase(txtLibraryName.Text) & ".GLB|" & _
                "was created in folder|" & UCase(txtPath.Text) & "|", _
                "i", , "Confirmation"
    End If
    
    m.bOK = True
    Me.Hide
    
ErrExit:
    vsStatus.Visible = False
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryPackager.Export", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrExport_Timer
'' Description: When the timer goes off, export the library
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrExport_Timer()
On Error GoTo ErrSection:

    tmrExport.Enabled = False
    Export

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryPackager.tmrExport.Timer", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnsureNoUserDeps
'' Description: Ensure that there are no user library dependencies
'' Inputs:      None
'' Returns:     True if no Dependencies, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EnsureNoUserDeps() As Boolean
On Error GoTo ErrSection:

    Dim astrDepends As New cGdArray     ' Array of user library dependencies
    Dim strMessage As String            ' Message to display to the user
    Dim lIndex As Long                  ' Index into a for loop
    Dim rs As Recordset                 ' Recordset into the database
    Dim strType As String               ' Type of the item
    Dim strID As String                 ' ID of the item
    Dim nSecurityLevel As Byte          ' Security level for the item
    Dim strPassword As String           ' Password for the item
    Dim astrDepends2 As cGdArray        ' Array of user library dependencies
    Dim lPos As Long                    ' Position in an array
    Dim strDependency As String         ' Dependency information
    
    EnsureNoUserDeps = True
    g.bReload = False
    
    nSecurityLevel = m.Library.SecurityLevel
    strPassword = m.Library.Password
    
    astrDepends.Create eGDARRAY_Strings
    LibraryDependencies astrDepends, m.Library.LibraryID, eGDDependencyFilter_UserLibraryOnly
    If astrDepends.Size > 0 Then
        Set astrDepends2 = New cGdArray
        For lIndex = 0 To astrDepends.Size - 1
            strDependency = Parse(astrDepends(lIndex), vbTab, 3) & vbTab & Parse(astrDepends(lIndex), vbTab, 2)
            If astrDepends2.BinarySearch(strDependency, lPos) = False Then
                astrDepends2.Add strDependency, lPos
            End If
        Next lIndex
        
        strType = ""
        strMessage = "Items in this library require the following item(s) in your User Library:|"
        For lIndex = 0 To astrDepends2.Size - 1
            If Parse(astrDepends2(lIndex), vbTab, 1) <> strType Then
                strType = Parse(astrDepends2(lIndex), vbTab, 1)
                strMessage = strMessage & "|" & strType & "(s):"
            End If
            strMessage = strMessage & "|" & Parse(astrDepends2(lIndex), vbTab, 2)
        Next lIndex
        strMessage = strMessage & "||Would you like to add these items to this library?"
        
        If InfBox(strMessage, "?", "+Add|-Cancel", "Library Export Warning") = "A" Then
            For lIndex = 0 To astrDepends.Size - 1
                strType = Parse(astrDepends(lIndex), vbTab, 3)
                strID = Parse(astrDepends(lIndex), vbTab, 1)
                
                Select Case UCase(strType)
                    Case "FUNCTION"
                        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] WHERE [FunctionID]=" & strID & ";", dbOpenDynaset)
                        If Not (rs.BOF And rs.EOF) Then
                            rs.Edit
                            rs!LibraryID = m.Library.LibraryID
                            rs!SecurityLevel = nSecurityLevel
                            EncryptField rs!Password, strPassword
                            rs!CheckSum = BuildCheckSum(rs, "tblFunctions")
                            rs.Update
                        End If
                    
                    Case "RULE"
                        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] WHERE [RuleID]=" & strID & ";", dbOpenDynaset)
                        If Not (rs.BOF And rs.EOF) Then
                            rs.Edit
                            rs!LibraryID = m.Library.LibraryID
                            rs!SecurityLevel = nSecurityLevel
                            EncryptField rs!Password, strPassword
                            rs!CheckSum = BuildCheckSum(rs, "tblRules")
                            rs.Update
                        End If
                        
                    Case "SYSTEM"
                        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] WHERE [SystemNumber]=" & strID & ";", dbOpenDynaset)
                        If Not (rs.BOF And rs.EOF) Then
                            rs.Edit
                            rs!LibraryID = m.Library.LibraryID
                            rs!SecurityLevel = nSecurityLevel
                            EncryptField rs!Password, strPassword
                            rs!CheckSum = BuildCheckSum(rs, "tblSystems")
                            rs.Update
                        End If
                        
                
                End Select
            Next lIndex
            g.bReload = True
        Else
            EnsureNoUserDeps = False
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLibraryPackager.EnsureNoUserDeps", eGDRaiseError_Raise, g.strAppPath
    
End Function
