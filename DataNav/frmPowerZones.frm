VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmPowerZones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ToTheTick PowerZones Login"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   1215
      Left            =   60
      TabIndex        =   1
      Top             =   1620
      Width           =   3735
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
      Caption         =   "frmPowerZones.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPowerZones.frx":0068
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPowerZones.frx":0088
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtPassword 
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Top             =   780
         Width           =   2415
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPowerZones.frx":00A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   0
         MultiLine       =   0   'False
         Alignment       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmPowerZones.frx":00C4
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPowerZones.frx":00E4
      End
      Begin HexUniControls.ctlUniTextBoxXP txtUserID 
         Height          =   315
         Left            =   900
         TabIndex        =   2
         Top             =   360
         Width           =   2655
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPowerZones.frx":0100
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   0
         MultiLine       =   0   'False
         Alignment       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmPowerZones.frx":0120
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPowerZones.frx":0140
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   255
         Left            =   180
         Top             =   780
         Width           =   915
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
         Caption         =   "frmPowerZones.frx":015C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPowerZones.frx":0190
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPowerZones.frx":01B0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   255
         Left            =   180
         Top             =   420
         Width           =   795
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
         Caption         =   "frmPowerZones.frx":01CC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPowerZones.frx":01FE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPowerZones.frx":021E
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   3960
      TabIndex        =   3
      Top             =   2340
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
      Caption         =   "frmPowerZones.frx":023A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmPowerZones.frx":0268
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmPowerZones.frx":0288
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Default         =   -1  'True
      Height          =   435
      Left            =   3960
      TabIndex        =   0
      Top             =   1740
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
      Caption         =   "frmPowerZones.frx":02A4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmPowerZones.frx":02CA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmPowerZones.frx":02EA
      RightToLeft     =   0   'False
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   60
      Picture         =   "frmPowerZones.frx":0306
      Stretch         =   -1  'True
      Top             =   60
      Width           =   5025
   End
End
Attribute VB_Name = "frmPowerZones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    bShown As Boolean
    bValid As Boolean
    strUserID As String
    strPassword As String
    
    PivotTables As cGdTree
    LastRefreshed As cGdTree
End Type
Private m As mPrivate

Private Sub cmdCancel_Click()

    m.bValid = False
    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim s$

    m.strUserID = Trim(txtUserID.Text)
    If UCase(m.strUserID) = "DEBUG" And IsIDE Then
        m.strUserID = "aamar@tickstrike.com" ' a valid user ID for our testing purposes
    End If
    
    ' password is not even really used, so we'll just change it to stars
    s = Trim(txtPassword.Text)
    If Len(s) > 0 And Left(s, 1) <> "*" Then
        m.strPassword = s
    End If
    
    SetIniFileProperty "UserID", m.strUserID, "PowerZones", g.strIniFile
    SetIniFileProperty "Password", m.strPassword, "PowerZones", g.strIniFile
    
    m.bValid = True
    Unload Me

End Sub

Private Sub Form_Load()

    Me.Icon = Picture16("kBlank")
    CenterTheForm Me
    
    g.Styler.StyleForm Me

End Sub

Private Sub txtPassword_Change()
    
    FixControls

End Sub

Private Sub txtUserID_Change()

    FixControls

End Sub

Private Sub FixControls()

    On Error Resume Next
    If Len(txtUserID.Text) > 0 And Len(txtPassword.Text) > 0 Then
        cmdOK.Enabled = True
        cmdOK.Default = True
    Else
        cmdOK.Enabled = False
    End If

End Sub

Public Function ZoneFileForSymbol(ByVal strSymbol$) As String

    Dim strZoneFile$
    
    ' determine the "zone file" for the symbol
    If IsForex(strSymbol) Then
        strZoneFile = StripStr(strSymbol, "$-")
    ElseIf SecurityType(strSymbol) = "F" Then
        strSymbol = PrimaryFutureBase(strSymbol)
        Select Case Parse(strSymbol, "-", 1)
        Case "SP", "ES"
            strZoneFile = "SandP500"
        Case "DJ", "YM"
            strZoneFile = "DowJones"
        Case "ND", "NQ"
            strZoneFile = "Nasdaq"
        Case "TF"
            strZoneFile = "Russell"
        Case "GC", "XK", "ZG", "QO"
            strZoneFile = "Gold"
        Case "CL", "QM"
            strZoneFile = "CrudeOil"
        Case "EU", "G6E"
            strZoneFile = "6e"
        Case "BP", "G6B"
            strZoneFile = "6b"
        Case "GX"
            strZoneFile = "fdax"
        Case "EX"
            strZoneFile = "fesx"
        End Select
    End If
    
    ZoneFileForSymbol = LCase(strZoneFile)

End Function

Public Function GetZoneData(ByVal strZoneFile$) As cGdTable

    Dim dRefreshed#
    Dim PivotTable As cGdTable

    ' just first time: show login form
    If Not m.bShown Then
        m.bShown = True
        m.bValid = False
        Set m.PivotTables = New cGdTree
        Set m.LastRefreshed = New cGdTree
        txtUserID.Text = GetIniFileProperty("UserID", m.strUserID, "PowerZones", g.strIniFile)
        m.strPassword = GetIniFileProperty("Password", m.strPassword, "PowerZones", g.strIniFile)
        If Len(m.strPassword) > 0 Then
            txtPassword.Text = String(Len(m.strPassword), "*")
        End If
        FixControls
        ShowForm Me, eForm_Modal
    End If
    
    ' if table was last loaded over 30 minutes ago, then reload it from their website
    strZoneFile = LCase(Trim(strZoneFile))
    If m.bValid And Len(strZoneFile) > 0 Then
        If m.LastRefreshed.Exists(strZoneFile) Then
            dRefreshed = Val(m.LastRefreshed(strZoneFile))
        End If
        If dRefreshed < Now - 30 / 1440# Then
            m.LastRefreshed(strZoneFile) = Str(CDbl(Now))
            Set PivotTable = LoadPivotTable(strZoneFile)
            ' but only replace the table if we got a good one from their website
            ' (e.g. in case of their website being down momentarily, etc)
            If PivotTable.NumRecords > 0 And PivotTable.NumFields > 3 Then
                Set m.PivotTables(strZoneFile) = PivotTable
            End If
        End If
    End If

    ' get table from the collection
    Set PivotTable = m.PivotTables(strZoneFile)
    If PivotTable Is Nothing Then
        ' if not exist, just return an empty table
        Set PivotTable = New cGdTable
    End If
    
    Set GetZoneData = PivotTable

End Function

'http://www.tothetick.com/fetchcsv.php?mkt=sandp500&user=aamar@tickstrike.com
'Store pivot data into gdTable (rows sorted by SessionDate, fields are prices in increasing sequence):
' SessionDate, From1, To1, From2, To2, ..., From12, To12
Private Function LoadPivotTable(ByVal strZoneFile$) As cGdTable

    Dim d#, dPrev#, iRec&, iZone&, iFirstGoodRec&, iLine&
    Dim s$, strWeb$
    Dim aLines As New cGdArray
    Dim aFields As New cGdArray
    Dim aData As New cGdTable
       
    ' append random number as arg just to override any browser web-page-caching
    strWeb = "http://www.tothetick.com/fetchcsv.php?mkt=" & strZoneFile & "&user=" & m.strUserID _
        & "&rand=" & Str(RandomNum(1, 9999))
    s = GetWebPageData(strWeb, 5)
    'FileFromString "c:\temp\Test1.txt", s
    s = Replace(s, vbCrLf, vbTab)
    s = Replace(s, vbLf, vbTab)
    s = Replace(s, vbCr, vbTab)
    aLines.SplitFields s, vbTab
    ' remove any blank lines
    For iLine = aLines.Size - 1 To 0 Step -1
        If Len(Trim(aLines(iLine))) = 0 Then
            aLines.Remove iLine
        End If
    Next
    
    aData.Clear
    If aLines.Size >= 10 Then
        ' first line is dates: 15/04/2011,18/04/2011,19/04/2011,...,06/06/2013,07/06/2013,10/06/2013
        ' store dates in first column of table
        aFields.SplitFields aLines(0), ","
        aData.CreateField eGDARRAY_Longs, 0
        aData.NumRecords = aFields.Size
        iFirstGoodRec = 0
        dPrev = 0
        For iRec = 0 To aFields.Size - 1
            s = Trim(aFields(iRec))
            If Len(s) = 10 And Mid(s, 3, 1) = "/" And Mid(s, 6, 1) = "/" Then
                ' convert from DD/MM/YYYY to YYYYMMDD, then to Julian
                s = Right(s, 4) & Mid(s, 4, 2) & Left(s, 2)
                d = DateOf(Val(s))
                If d < 25000 Then
                    iFirstGoodRec = iRec + 1 ' bad date
                Else
                    aData.Num(0, iRec) = d
                    If d <= dPrev Then
                        iFirstGoodRec = iRec  ' date of last record must have been bad
                    End If
                    dPrev = d
                End If
            End If
        Next
        For iRec = 0 To iFirstGoodRec - 1
            aData.Num(0, iRec) = kNullData
        Next
    
        ' now parse out the values for each zone (table fields) at each date (table record)
        iZone = 0
        For iLine = aLines.Size - 1 To 1 Step -1
            s = aLines(iLine)
            If Len(s) > 10 And InStr(s, ",") > 1 Then
                ' add 2 columns for each zone
                iZone = iZone + 1
                aFields.SplitFields s, ","
                aData.CreateField eGDARRAY_Doubles, , "From" & Str(iZone)
                aData.CreateField eGDARRAY_Doubles, , "To" & Str(iZone)
                For iRec = iFirstGoodRec To aFields.Size - 1
                    s = aFields(iRec)
                    If InStr(s, "-") > 1 Then
                        d = Val(Parse(s, "-", 1))
                        aData.Num(aData.NumFields - 2, iRec) = d
                        d = Val(Parse(s, "-", 2))
                        aData.Num(aData.NumFields - 1, iRec) = d
                    End If
                Next
            End If
        Next
    End If
    
    If 0 Then
        s = aData.ToString
        FileFromString "c:\temp\Test2.txt", s
    End If
    Set LoadPivotTable = aData

End Function



