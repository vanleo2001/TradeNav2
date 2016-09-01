VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "VSOCX6.OCX"
Begin VB.Form frmLibrary 
   Caption         =   "Library Editor"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   8730
   Begin VB.CommandButton cmdCorner 
      Caption         =   "Corner"
      Height          =   375
      Left            =   7440
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4305
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtLibraryID 
      Height          =   285
      Left            =   6960
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "txtLibraryID"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   420
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1200
   End
   Begin vsOcx6LibCtl.vsIndexTab vsitLibrary 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6800
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
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483635
      Caption         =   "Se&ttings|&Advanced"
      Align           =   0
      Appearance      =   1
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   -1  'True
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      Begin VB.Frame fraAdvanced 
         BorderStyle     =   0  'None
         Height          =   3480
         Left            =   9060
         TabIndex        =   9
         Top             =   330
         Width           =   8325
         Begin VB.Frame fraSecurity 
            Caption         =   "Security"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   360
            TabIndex        =   15
            Top             =   360
            Width           =   4935
            Begin VB.TextBox txtPassword 
               Height          =   315
               Left            =   1365
               TabIndex        =   4
               Text            =   "txtPassword"
               Top             =   480
               Width           =   3120
            End
            Begin VB.ComboBox cboSecurityLevel 
               Height          =   315
               Left            =   1365
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   840
               Width           =   3120
            End
            Begin VB.CheckBox chkCannotDelete 
               Caption         =   "This rule cannot be deleted"
               Height          =   255
               Left            =   1365
               TabIndex        =   6
               Top             =   1320
               Width           =   2535
            End
            Begin VB.Label Label7 
               Caption         =   "Password"
               Height          =   255
               Left            =   240
               TabIndex        =   17
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label8 
               Caption         =   "Security Level"
               Height          =   255
               Left            =   240
               TabIndex        =   16
               Top             =   840
               Width           =   1095
            End
         End
      End
      Begin VB.Frame fraSettings 
         BorderStyle     =   0  'None
         Height          =   3480
         Left            =   45
         TabIndex        =   8
         Top             =   330
         Width           =   8325
         Begin VB.TextBox txtDescription 
            Height          =   1095
            Left            =   1920
            TabIndex        =   3
            Text            =   "txtDescription"
            Top             =   1680
            Width           =   4935
         End
         Begin VB.TextBox txtAuthor 
            Height          =   285
            Left            =   1920
            TabIndex        =   2
            Text            =   "txtAuthor"
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox txtLastModified 
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "txtLastModified"
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   1920
            TabIndex        =   1
            Text            =   "txtName"
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label4 
            Caption         =   "Description"
            Height          =   255
            Left            =   480
            TabIndex        =   13
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Author"
            Height          =   255
            Left            =   480
            TabIndex        =   12
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Last Modified"
            Height          =   255
            Left            =   480
            TabIndex        =   11
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Library Name"
            Height          =   255
            Left            =   480
            TabIndex        =   10
            Top             =   600
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frmLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private gbDirty As Boolean
Private gstrName As String
Private gstrPassword As String
Private glSecurityLevel As Long

Public Sub LoadRec(ByVal lLibraryID As Long)
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    Dim RetVal As Long
        
    RetVal = LockWindowUpdate(Me.HWND)
    
    Set rs = dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] WHERE [LibraryID] = " & _
        str(lLibraryID) & " ORDER BY [LibraryID];", dbOpenDynaset)
    
    With rs
        If Not IsNull(!LibraryID) Then txtLibraryID.Text = !LibraryID Else txtLibraryID.Text = ""
        If Not IsNull(!LibraryName) Then txtName.Text = !LibraryName Else txtName.Text = ""
        gstrName = txtName.Text
        If InStr(Me.Caption, " - ") = 0 Then
            Me.Caption = Me.Caption + " - " + txtName.Text
        Else
            Me.Caption = Mid(Me.Caption, 1, InStr(Me.Caption, " - ") - 1) + " - " + txtName.Text
        End If
        If Not IsNull(!LastModified) Then txtLastModified.Text = Format(!LastModified, "mmm d, yyyy h:mm AMPM") Else txtLastModified.Text = ""
        If Not IsNull(!Author) Then txtAuthor.Text = !Author Else txtAuthor.Text = ""
        If Not IsNull(!LibraryDesc) Then txtDescription.Text = !LibraryDesc Else txtDescription.Text = ""
        
        If Not IsNull(!SecurityLevel) Then cboSecurityLevel.ListIndex = !SecurityLevel Else cboSecurityLevel.ListIndex = EDITANDVIEW
        glSecurityLevel = cboSecurityLevel.ListIndex
        If Not IsNull(!Password) Then txtPassword.Text = !Password Else txtPassword.Text = ""
        gstrPassword = txtPassword.Text
        If (!CannotDelete = True) Then
            chkCannotDelete.value = vbChecked
        Else
            chkCannotDelete.value = vbUnchecked
        End If
    End With
        
    SetDirtyFlag False
    
ErrExit:
    Set rs = Nothing
    RetVal = LockWindowUpdate(0)
    Exit Sub
ErrSection:
    ShowMsg
    Resume ErrExit
End Sub

Private Sub SetDirtyFlag(ByVal bValue As Boolean)
    
    gbDirty = bValue
    cmdSave.Enabled = bValue

End Sub

Public Sub Add()
    
    txtLibraryID.Text = "-1"
    txtName.Text = ""
    gstrName = ""
    txtLastModified.Text = ""
    txtAuthor.Text = ""
    txtDescription.Text = ""
    
    cboSecurityLevel.ListIndex = EDITANDVIEW
    glSecurityLevel = EDITANDVIEW
    txtPassword.Text = ""
    gstrPassword = txtPassword.Text
    chkCannotDelete.value = vbUnchecked
    
    SetDirtyFlag False

End Sub

Private Sub cboSecurityLevel_Change()
    SetDirtyFlag True
End Sub

Private Sub cboSecurityLevel_Click()
    SetDirtyFlag True
End Sub

Private Sub chkCannotDelete_Click()
    SetDirtyFlag True
End Sub

Private Sub cmdSave_Click()
    Save
End Sub

Private Sub Form_Activate()
    
    SetDirtyFlag False

End Sub

Private Sub Form_Load()
    Dim lPropVal As Long
    
    'Mod:
    Me.Icon = frmMain.ImageList1.ListImages(6).ExtractIcon
    
    ReSizeMDIChildForm Me, cmdCorner
    CenterTheForm Me
        
    lPropVal = IniFileProperty("LibraryEditHeight", 0, False, "Placement", App.Path + "\NAVDEFS.INI")
    If lPropVal <> 0 Then Me.Height = lPropVal
    
    cboSecurityLevel.AddItem "Can Edit", 0
    cboSecurityLevel.AddItem "Can View but not Edit", 1
    cboSecurityLevel.AddItem "Cannot View", 2
    If gbGenesisUser = True Then cboSecurityLevel.AddItem "Will not show in list", 3
    chkCannotDelete.Visible = gbGenesisUser
    txtPassword.PasswordChar = "*"
    
    SetDirtyFlag False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim iResult%
        
    If gbDirty Then
        iResult = MsgBox("Library has changed.  Do you want to save the changes?", vbQuestion + vbYesNoCancel, "Confirmation")
        If iResult = vbCancel Then
            Cancel = True
            Exit Sub
        ElseIf iResult = vbYes Then
            cmdSave_Click
        End If
    End If
    
    gbDirty = False
    cmdSave.Enabled = False
    
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
    
    'Mod:
    frmSelect.Show
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    IniFileProperty "LibraryEditTop", Me.top, True, "Placement", App.Path + "\NAVDEFS.INI"
    IniFileProperty "LibraryEditHeight", Me.Height, True, "Placement", App.Path + "\NAVDEFS.INI"
    
End Sub

Private Sub txtAuthor_Change()
    SetDirtyFlag True
End Sub

Private Sub txtDescription_Change()
    SetDirtyFlag True
End Sub

Private Sub txtName_Change()
    SetDirtyFlag True
End Sub

Private Sub txtPassword_Change()
    SetDirtyFlag True
End Sub

Private Sub Save()
On Error GoTo ErrSection:
    
    Dim QryDef As QueryDef
    Dim rs As Recordset, rs2 As Recordset
    Dim RetVal As Long
    Dim bNewRec As Boolean
    
    txtName.Text = Trim(txtName.Text)
    If Len(gstrName) > 0 Then
        If txtName.Text <> gstrName Then
            If LibraryExists(txtName.Text) = False Then
                If RenameOrCopy("library") <> "R" Then txtLibraryID.Text = "-1"
            Else
                MsgBox "Library " & txtName.Text & " already exists", vbExclamation, "Warning"
                GoTo ErrExit
            End If
        End If
    End If
    
    RetVal = LockWindowUpdate(Me.HWND)
    
    bNewRec = False
    Set rs = dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] WHERE [LibraryID] = " & _
        txtLibraryID.Text & " ORDER BY [LibraryID];", dbOpenDynaset)
    
    With rs
        If rs.EOF Then
            bNewRec = True
            Set QryDef = dbNav.QueryDefs("qryMaxLibraryID")
            Set rs2 = QryDef.OpenRecordset()
            .AddNew
            !LibraryID = rs2!LibraryID + 1
        Else
            .Edit
        End If
        
        !LibraryName = txtName.Text
        !LibraryDesc = txtDescription.Text
        !Author = txtAuthor.Text
        !SecurityLevel = cboSecurityLevel.ListIndex
        !Password = txtPassword.Text
        !CannotDelete = chkCannotDelete.value
        
        !Path = ""
        !version = ""
        !CustomerID = 0
        
        .Update
        
        If bNewRec = True Then
            .MoveLast
            txtLibraryID.Text = str(!LibraryID)
        End If
        
        .Edit
        !LastModified = Now()
        .Update
        
    End With
    
    SetDirtyFlag False
    gReloadMenu = True

ErrExit:
    Set rs = Nothing
    Set rs2 = Nothing
    Set QryDef = Nothing
    RetVal = LockWindowUpdate(0)
    Exit Sub
ErrSection:
    ShowMsg
    Resume ErrExit
End Sub

Private Function LibraryExists(strLibraryName As String) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim QryDef As QueryDef
    
    LibraryExists = False
    Set QryDef = dbNav.QueryDefs("qryLibraryIDFromName")
    QryDef.Parameters(0).value = strLibraryName
    Set rs = QryDef.OpenRecordset
    
    If rs.RecordCount = 0 Then
        LibraryExists = False
    Else
        LibraryExists = True
    End If
        
ErrExit:
    Set rs = Nothing
    Set QryDef = Nothing
    Exit Function
ErrSection:
    ShowMsg
    Resume ErrExit:
End Function

