VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTest2 
   Caption         =   "Test 2"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   5640
   Begin VB.Timer tmrDownload 
      Left            =   3720
      Top             =   4980
   End
   Begin VB.Timer tmrTest2 
      Left            =   4680
      Top             =   4980
   End
   Begin VB.Timer tmrTurnkey 
      Left            =   5160
      Top             =   4980
   End
   Begin VB.Timer tmrTest1 
      Left            =   4200
      Top             =   4980
   End
   Begin HexUniControls.ctlUniListBoxXP lst 
      Height          =   3570
      Left            =   1800
      TabIndex        =   9
      Top             =   300
      Width           =   2955
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
      TrapTab         =   0   'False
      Tip             =   "Frmtest2.frx":0000
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "Frmtest2.frx":0020
      ManualStart     =   0   'False
      Columns         =   0
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtCallback 
      Height          =   315
      Left            =   3660
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "Frmtest2.frx":003C
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
      Tip             =   "Frmtest2.frx":0066
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "Frmtest2.frx":0086
   End
   Begin HexUniControls.ctlUniFrameWL Frame1 
      Height          =   5235
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
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
      Caption         =   "Frmtest2.frx":00A2
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "Frmtest2.frx":00CE
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "Frmtest2.frx":00EE
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP Command7 
         Height          =   375
         Left            =   180
         TabIndex        =   7
         Top             =   3642
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
         Caption         =   "Frmtest2.frx":010A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Frmtest2.frx":013E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Frmtest2.frx":015E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command6 
         Height          =   375
         Left            =   180
         TabIndex        =   6
         Top             =   3085
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
         Caption         =   "Frmtest2.frx":017A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Frmtest2.frx":01AE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Frmtest2.frx":01CE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command5 
         Height          =   375
         Left            =   180
         TabIndex        =   5
         Top             =   2528
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
         Caption         =   "Frmtest2.frx":01EA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Frmtest2.frx":021E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Frmtest2.frx":023E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command4 
         Height          =   375
         Left            =   180
         TabIndex        =   4
         Top             =   1971
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
         Caption         =   "Frmtest2.frx":025A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Frmtest2.frx":028E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Frmtest2.frx":02AE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command3 
         Height          =   375
         Left            =   180
         TabIndex        =   3
         Top             =   1414
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
         Caption         =   "Frmtest2.frx":02CA
         BackColor       =   -2147483639
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Frmtest2.frx":02FE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Frmtest2.frx":031E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command2 
         Height          =   375
         Left            =   180
         TabIndex        =   2
         Top             =   857
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
         Caption         =   "Frmtest2.frx":033A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Frmtest2.frx":036E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Frmtest2.frx":038E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP Command1 
         Height          =   375
         Left            =   180
         TabIndex        =   1
         Top             =   300
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
         Caption         =   "Frmtest2.frx":03AA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "Frmtest2.frx":03DE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "Frmtest2.frx":03FE
         RightToLeft     =   0   'False
      End
      Begin NavSuite.gdPriceEditor pePrice 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   4200
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
      End
   End
End
Attribute VB_Name = "frmTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTest2.frm
'' Description: Test form for development
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 05/18/2010   DAJ         Added the future table analyzer
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strBrokerName As String
    
    dOnCloseTimeExch As Double
    
    Bars As cGdBars
    strTradeSymbol As String
    
    MjkToGen As cGdTree
    GenToMjk As cGdTree
    
    bIgnoreAppLoaded As Boolean
    Downloader As cDownloader
End Type
Private m As mPrivate

Public DoClearColors As Boolean

Private Property Get TradeSymbol() As String
    TradeSymbol = m.strTradeSymbol
End Property
Private Property Let TradeSymbol(ByVal strTradeSymbol As String)
    m.strTradeSymbol = strTradeSymbol
End Property

Public Property Get IgnoreAppLoaded() As Boolean
    IgnoreAppLoaded = m.bIgnoreAppLoaded
End Property

Private Sub Command1_Click()
On Error GoTo ErrSection:

    DownloaderSample
    'g.IntBroker.GetSymbolAvailability ""

#If 0 Then
    Dim lIndex As Long
    Dim strSymbol As String
    Dim Bars As cGdBars
    Dim dEndTime As Double
    Dim strEndTime As String
    Dim astrTimes As cGdArray
    Dim lPos As Long
    
    Set astrTimes = New cGdArray
    astrTimes.Create eGDARRAY_Strings

    For lIndex = 1 To g.SymbolPool.NumRecords
        If g.SymbolPool.SecType(lIndex) = eSYMType_Future Then
            strSymbol = g.SymbolPool.Symbol(lIndex)
            
            If InStr(strSymbol, "-067") <> 0 Then
                Set Bars = New cGdBars
                SetBarProperties Bars, strSymbol
                
                dEndTime = ConvertTimeZone(Date + (Bars.Prop(eBARS_EndTime) / 1440#), Bars.Prop(eBARS_ExchangeTimeZoneInf), "")
                strEndTime = Format(dEndTime, "00000.000000")
                
                astrTimes.BinarySearch strEndTime & vbTab, lPos, eGdSort_MatchUsingSearchStringLength
                astrTimes.Add strEndTime & vbTab & strSymbol, lPos
            End If
        End If
    Next lIndex
    
    astrTimes.ToFile AddSlash(App.Path) & "EndTimes.TXT"
    
    AddList "Done"
#End If
    
#If 0 Then
    Dim lIndex As Long
    Dim strSymbol As String
    Dim strBaseSymbol As String
    Dim strContract As String
    Dim astrBaseSymbols As cGdArray
    Dim lPos As Long
    Dim Rules As cGdTree
    Dim astrFile As cGdArray
    Dim bAdd As Boolean
    Dim Bars As cGdBars
    
    Set astrFile = New cGdArray
    If astrFile.FromFile("N:\Mc1\Cfg\FutExpRules.TXT") Then
        LoadGenToMjk
        
        Set astrBaseSymbols = New cGdArray
        astrBaseSymbols.Create eGDARRAY_Strings
        
        Set Rules = New cGdTree
        For lIndex = 0 To astrFile.Size - 1
            strSymbol = Parse(astrFile(lIndex), vbTab, 1)
            If Rules.Exists(strSymbol) = False Then
                Rules.Add Parse(astrFile(lIndex), vbTab, 3), strSymbol
            End If
        Next lIndex
        
        For lIndex = 1 To g.SymbolPool.NumRecords
            If g.SymbolPool.SecType(lIndex) = eSYMType_Future Then
                strSymbol = g.SymbolPool.Symbol(lIndex)
                
                strBaseSymbol = Parse(strSymbol, "-", 1)
                strContract = Parse(strSymbol, "-", 2)
                
                If Len(strContract) > 4 Then
                    If Left(strContract, 4) = "2014" Then
                        bAdd = Not Rules.Exists(strBaseSymbol)
                        If bAdd = False Then
                            bAdd = (Len(Rules(strBaseSymbol)) = 0)
                        End If
                        
                        If bAdd = True Then
                            If astrBaseSymbols.BinarySearch(strBaseSymbol & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) = False Then
                                astrBaseSymbols.Add strBaseSymbol & vbTab & m.GenToMjk(strBaseSymbol), lPos
                            End If
                        
                            AddList g.SymbolPool.Symbol(lIndex)
                        End If
                    End If
                End If
            End If
        Next lIndex
        
        astrBaseSymbols.ToFile AddSlash(App.Path) & "CurrentSymbols.TXT"
    End If
#End If

#If 0 Then
    Dim dDate1 As Double
    Dim dDate2 As Double
    Dim lIndex As Long
    
    For lIndex = 1 To 12
        'dDate1 = GetDateFromRule(2014, lIndex, "10")
        'dDate2 = GetDateFromRule(2014, lIndex, "FB+9B")
        
        'AddList DateFormat(dDate1, MM_DD_YYYY) & " ; " & DateFormat(dDate2, MM_DD_YYYY)
        
        dDate1 = GetDateFromRule(2014, lIndex, "3FN-30")
        AddList DateFormat(dDate1, MM_DD_YYYY)
    Next lIndex
#End If

#If 0 Then
    Dim astrFile As cGdArray
    Dim astrLine As cGdArray
    Dim lIndex As Long
    Dim strBase As String
    Dim strContract As String
    Dim strMjkBase As String
    
    LoadMjkToGen
    
    Set astrFile = New cGdArray
    If astrFile.FromFile("N:\Mc1\Cfg\FutExpRules.TXT") Then
        For lIndex = 0 To astrFile.Size - 1
            Set astrLine = New cGdArray
            astrLine.SplitFields astrFile(lIndex), "="
            
            strBase = ""
            If m.MjkToGen.Exists(astrLine(0)) Then
                strBase = m.MjkToGen(astrLine(0))
            End If
            
            astrFile(lIndex) = strBase & vbTab & astrLine(0) & vbTab & astrLine(1)
        Next lIndex
        
        astrFile.ToFile "N:\Mc1\Cfg\FutExpRules2.TXT"
    End If
    
    AddList "Done"
#End If

    'g.Broker.ReplayLogFile "S:\Pfeiffer\CQG20140514.log", eTT_AccountType_CQG
    'g.Broker.ReplayLogFile "S:\Pfeiffer\RJO20140514.log", eTT_AccountType_RjoCqg
    'g.Broker.ReplayLogFile "S:\Pfeiffer\RJO20140515.log", eTT_AccountType_RjoCqg
    'g.Broker.ReplayLogFile "S:\Pfeiffer\CQG20140515.log", eTT_AccountType_CQG

    'KeepOnlyMessagesReceived "S:\Pfeiffer\TN20140514.log", eTT_AccountType_CQG, "S:\Pfeiffer\20140514.REC"
    'KeepOnlyMessagesReceived "S:\Pfeiffer\TN20140515.log", eTT_AccountType_CQG, "S:\Pfeiffer\20140515.REC"

    'SplitBigAscFile "S:\FX0430_I.ASC", AddSlash(App.Path) & "IbFxGen"
    'AscFileCompare "S:\FX0430_I.ASC", AddSlash(App.Path) & "IbFx", "IbFxCompare.TXT"
    'MergeAscFiles AddSlash(App.Path) & "IbForex.TXT", AddSlash(App.Path) & "IbFxGen", AddSlash(App.Path) & "IbFx", AddSlash(App.Path) & "IbFxOut"
    
    'Dim lPeriod As Long
    
    'lPeriod = GetPeriodicity("FractZen")
    'AddList "FractZen Periodicity = " & Str(lPeriod) & "; Daily = " & Str(ePRD_Days) & "; " & Str(lPeriod >= (ePRD_Days + 1))
    
    'AddList KillProcess("GenRjoCqg", False)

#If 0 Then
    Dim astrSymbols As cGdArray         ' Array of symbols from IB object
    Dim astrOutputFile As cGdArray      ' Output file
    Dim lIndex As Long                  ' Index into a for loop
    
    If g.IntBroker Is Nothing Then
        AddList "g.IntBroker is Nothing"
    Else
        If 0 Then
            Set astrSymbols = g.IntBroker.GenesisSymbolList(False)
            Set astrOutputFile = New cGdArray
            astrOutputFile.Create eGDARRAY_Strings
            
            For lIndex = 0 To astrSymbols.Size - 1
                If InStr(astrSymbols(lIndex), "@") <> 0 Then
                    astrOutputFile.Add astrSymbols(lIndex)
                    AddList astrSymbols(lIndex)
                End If
            Next lIndex
            
            astrOutputFile.ToFile AddSlash(App.Path) & "IbForex.TXT"
            AddList "Done"
        Else
            frmIbDataPull.ShowMe
        End If
    End If
#End If

'    Dim BInfo As cBrokerInfo
'    Dim lIndex As Long
'
'    Set BInfo = g.Broker.BrokerInfo(eTT_AccountType_Oec)
'    If Not BInfo Is Nothing Then
'        For lIndex = 1 To BInfo.Accounts.Count
'            With BInfo.Accounts(lIndex)
'                AddList .AccountNumber & " --> " & .UserName & "; " & .Password
'            End With
'        Next lIndex
'    End If
'
'    Dim lIndex As Long
'
'    For lIndex = JulFromLong(20140201) To JulFromLong(20140219)
'
'        AddList "ZB-057;8;3;" & DateFormat(lIndex, MM_DD_YYYY) & " = " & Str(CalcAutoBreakoutRange("ZB-057", 8, 3, lIndex))
'        AddList "ZB-057;6;6;" & DateFormat(lIndex, MM_DD_YYYY) & " = " & Str(CalcAutoBreakoutRange("ZB-057", 6, 6, lIndex))
'
'    Next lIndex

    'Dim Ration As cBrokerMessage
    
    'Set Ration = New cBrokerMessage
    'frmTurnkeyManage.ShowMeRation Ration
    'frmTurnkeyEditor.ShowMeRation Ration

'    Dim Bars As cGdBars
'
'    Set Bars = New cGdBars
'
'    DM_GetBars Bars, "ES-067", "5 minute", -624
'    Bars.ToFile "GDB", AddSlash(App.Path) & "ES.DAJ"
'    AddList "ES - " & DateFormat(Bars(eBARS_DateTime, 0), MM_DD_YYYY, HH_MM_SS)
'
'    DM_GetBars Bars, "AP-067", "5 minute", -624
'    AddList "AP - " & DateFormat(Bars(eBARS_DateTime, 0), MM_DD_YYYY, HH_MM_SS)

    'AddList "FptOec User = " & g.Broker.IsBrokerUser(eTT_AccountType_FptOec)
    'AddList "FptOec Object Is Nothing = " & Str(g.FptOec Is Nothing)

    'FixTotVolAndOi JulFromLong(20090630), JulFromLong(20131016)
    'FixTotVolAndOi JulFromLong(19000101), JulFromLong(20131031), ",CL,CL2,CL3,HO,HO2,HO3,IC,NG,NG2,NG3,RB,RB2,RB3,"
    'ContractOiDiffs "CL3", "CL", "CL2"
    
    'ReadNifty "S:\MjkBak\MJK\V090722.N16", "S:\MjkBak\Asc"
    
    'AddList g.Profit.Profit("LE-201313 P130.0", 3.75, 7)
    'AddList g.Profit.Profit("LE-201313 C136.0", 2.1, 7)
    
'    Dim lNumber As Long
'
'    lNumber = 1
'    AddList Str(lNumber) & " AND 1 = " & Str(lNumber And 1)
'    AddList Str(lNumber) & " AND 2 = " & Str(lNumber And 2)
'    AddList Str(lNumber) & " BIT 1 = " & Str(GetBit(lNumber, 1))
'    AddList Str(lNumber) & " BIT 2 = " & Str(GetBit(lNumber, 2))
'
'    lNumber = 2
'    AddList Str(lNumber) & " AND 1 = " & Str(lNumber And 1)
'    AddList Str(lNumber) & " AND 2 = " & Str(lNumber And 2)
'    AddList Str(lNumber) & " BIT 1 = " & Str(GetBit(lNumber, 1))
'    AddList Str(lNumber) & " BIT 2 = " & Str(GetBit(lNumber, 2))
    
#If 0 Then
    Dim strMessage As String
    Dim Lot As cBrokerMessage
    Dim LotDetails As cGdTree
    
    'strMessage = "ID=1" & vbTab & "CategoryName=Lot Information" & vbTab & "TabNumber=1"
    'g.Turnkey.TurnkeyMessageReceived eGDTurnkeyMessage_LotColumnCategories, strMessage
    
    'strMessage = "ID=2" & vbTab & "CategoryName=Cattle Information" & vbTab & "TabNumber=2"
    'g.Turnkey.TurnkeyMessageReceived eGDTurnkeyMessage_LotColumnCategories, strMessage
    
    'strMessage = "ID=3" & vbTab & "CategoryName=Credits" & vbTab & "TabNumber=3"
    'g.Turnkey.TurnkeyMessageReceived eGDTurnkeyMessage_LotColumnCategories, strMessage
    
    'strMessage = "ID=4" & vbTab & "CategoryName=Debits" & vbTab & "TabNumber=4"
    'g.Turnkey.TurnkeyMessageReceived eGDTurnkeyMessage_LotColumnCategories, strMessage

    Set Lot = New cBrokerMessage
    Set LotDetails = New cGdTree

    frmTurnkeyEditLot.ShowMe Lot, LotDetails
#End If

#If 0 Then
    Dim astrFiles As cGdArray
    Dim astrFile As cGdArray
    Dim lIndex As Long
    Dim lIndex2 As Long
    
    
    Set astrFiles = New cGdArray
    astrFiles.GetMatchingFiles "S:\TotVol\*.DAT"
    
    For lIndex = 0 To astrFiles.Size - 1
        AddList astrFiles(lIndex)
        
        Set astrFile = New cGdArgs
        If astrFile.FromFile(astrFiles(lIndex)) Then
            For lIndex2 = 0 To astrFile.Size - 1
                If IsDigit(Left(astrFile(lIndex2), 1)) = False Then
                    astrFile(lIndex2) = "#" & astrFile(lIndex2)
                    AddList astrFile(lIndex2)
                End If
            Next lIndex2
            
            astrFile.ToFile astrFiles(lIndex)
        End If
    Next lIndex
    
    AddList "Done"
#End If

#If 0 Then
    Dim lDate As Long
    Dim lDate2 As Long
    Dim lYear As Long
    Dim lMonth As Long
    Dim strToAdd As String
    
    For lYear = 2000 To 2013
        For lMonth = 1 To 12
            'lDate = GetDateFromRule(lYear, lMonth, "F-15")
            'lDate2 = GetDateFromRule(lYear, lMonth, "F-11B")
            'lDate = GetDateFromRule(lYear, lMonth, "14-2B")
            'lDate2 = lDate
            
            lDate = GetDateFromRule(lYear, lMonth, "F-15")
            If IsWeekday(lDate) Then
                lDate = lDate - 1
                Do While IsWeekday(lDate) = False
                    lDate = lDate - 1
                Loop
            Else
                Do While Weekday(lDate) <> vbThursday
                    lDate = lDate - 1
                Loop
            End If
            lDate2 = GetDateFromRule(lYear, lMonth, "F-12B")
            
            If lDate = lDate2 Then
                strToAdd = ""
            Else
                strToAdd = "********"
            End If
            
            AddList Format(lMonth, "00") & "/" & Str(lYear) & " --> " & DateFormat(lDate, MM_DD_YYYY) & " ( " & DateFormat(lDate2, MM_DD_YYYY) & " ) " & strToAdd
        Next lMonth
    Next lYear
#End If

#If 0 Then
    Dim astrInput As cGdArray
    Dim astrOutput As cGdArray
    Dim lIndex As Long
    Dim strLine As String
    
    Set astrInput = New cGdArray
    Set astrOutput = New cGdArray
    astrOutput.Create eGDARRAY_Strings
    
    If astrInput.FromFile(AddSlash(App.Path) & "Provided\RateMap.TXT") Then
        For lIndex = 0 To astrInput.Size - 1
            strLine = astrInput(lIndex)
            
            If InStr(strLine, "@GAIN") <> 0 Then
                strLine = Replace(strLine, "@GAIN" & vbTab & "10000" & vbTab, "@GAIN" & vbTab & "100000" & vbTab)
            End If
            astrOutput.Add strLine
        Next lIndex
        
        astrOutput.ToFile AddSlash(App.Path) & "Provided\RateMap.DAJ"
    End If
    
    AddList "Done"
#End If

#If 0 Then
    Dim strSymbols As String
    Dim astrSymbols As cGdArray
    Dim lIndex As Long
    
    strSymbols = "OAUD,OCAD,OCHF,OCL,OED,OEN,OES,OEUR,OGBP,OGC,OHG,OHO,OJPY,ONG,ORB,OSI,OZB,OZC,OZF,OZL,OZM,OZN,OZS,OZT,OZW"
    
    Set astrSymbols = New cGdArray
    astrSymbols.SplitFields strSymbols, ","
    
    For lIndex = 0 To astrSymbols.Size - 1
        g.Oec.Broker.SendMessage eGDBrokerMessageType_GetTifs, astrSymbols(lIndex)
    Next lIndex
#End If

#If 0 Then
    Dim astrInputFile As cGdArray
    Dim astrOutputFile As cGdArray
    Dim lIndex As Long
    Dim astrFields As cGdArray
    Dim Cfs As cGdTree
    
    Set astrInputFile = New cGdArray
    Set astrOutputFile = New cGdArray
    astrOutputFile.Create eGDARRAY_Strings
    Set Cfs = New cGdTree
    
    If astrInputFile.FromFile("C:\Dvlp\Parse32\MjkSyms.DAJ") Then
        For lIndex = 0 To astrInputFile.Size - 1
            Set astrFields = New cGdArray
            astrFields.SplitFields astrInputFile(lIndex), ";"
            
            Cfs.Add astrFields(1), astrFields(0)
        Next lIndex
        
        For lIndex = 0 To astrInputFile.Size - 1
            Set astrFields = New cGdArray
            astrFields.SplitFields astrInputFile(lIndex), ";"
            
            If astrFields(2) <> astrFields(0) Then
                If Cfs.Exists(astrFields(2)) Then
                    If Cfs(astrFields(2)) <> astrFields(1) Then
                        astrOutputFile.Add astrInputFile(lIndex) & ";" & Cfs(astrFields(2))
                    End If
                End If
            End If
        Next lIndex
    End If
    
    astrOutputFile.ToFile AddSlash(App.Path) & "Diffs.DAJ"
#End If
              
#If 0 Then
    Dim astrFile As cGdArray            ' Input file
    Dim lIndex As Long                  ' Index into a for loop
    Dim Bars As cGdBars                 ' Bars object
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim astrFields As cGdArray          ' Fields for the line
    Dim lDate As Long                   ' Date
              
    LoadMjkToGen
    
    Set astrFile = New cGdArray
    If astrFile.FromFile("C:\Dvlp\Parse32\20130830.TOT") Then
        For lIndex = 0 To astrFile.Size - 1
            Set astrFields = New cGdArray
            astrFields.SplitFields astrFile(lIndex), vbTab
            
            lDate = CLng(Val(astrFields(0)))
            If m.MjkToGen.Exists(astrFields(1)) Then
                strGenesisSymbol = m.MjkToGen(astrFields(1))
                
                Set Bars = New cGdBars
                If DM_GetBars(Bars, strGenesisSymbol & "-057", "Daily", lDate) Then
                    If Bars(eBARS_DateTime, 0) = JulFromLong(lDate) Then
                        If Bars(eBARS_Vol, 0) <> CLng(Val(astrFields(2))) Then
                            AddList astrFields(1) & "; " & strGenesisSymbol & "; Volume " & Str(astrFields(2)) & " <> " & Str(Bars(eBARS_Vol, 0))
                        End If
                        If Bars(eBARS_OI, 0) <> CLng(Val(astrFields(3))) Then
                            AddList astrFields(1) & "; " & strGenesisSymbol & "; OI " & Str(astrFields(3)) & " <> " & Str(Bars(eBARS_OI, 0))
                        End If
                    End If
                End If
            Else
                AddList astrFields(1) & " could not be translated"
            End If
        Next lIndex
    End If
#End If
              
'    CombinedRollFiles
              
'    Dim strNumBars As String
'    Dim lNumBars As Long
'
'    strNumBars = "0|1|0"
'
'    If Len(strNumBars) > 0 Then
'        If CBool(Parse(strNumBars, "|", 1)) = True Then
'            AddList "First field is True"
'            lNumBars = CLng(Val(Parse(strNumBars, "|", 3))) * -1&
'        Else
'            AddList "First field is False"
'            lNumBars = CLng(Val(Parse(strNumBars, "|", 2))) * -1&
'        End If
'    Else
'        lNumBars = 0&
'    End If
'
'    AddList "Num Bars = " & Str(lNumBars)
              
    'g.Vision.GetTradeRoutes
              
    'g.Transact.Reconnect
              
    'AddList "AllowEnable=" & EncryptToHex("3,30,80,16262")
    'AddList "SmiZx=" & EncryptToHex("YC-,YC2-,YC3-")
    
    'SearchInProjectFiles "_Timer("
    'SearchInProjectFiles "gdResetProfiles"
    
#If 0 Then
    Dim lIndex As Long
    Dim strFile As String
    Dim astrSymbols As cGdArray
    
    Set astrSymbols = New cGdArray
    strFile = FileToString("C:\David\Shared\Symbols.TXT")
    astrSymbols.SplitFields strFile, vbTab
    
    For lIndex = 0 To astrSymbols.Size - 1
        astrSymbols(lIndex) = Parse(astrSymbols(lIndex), "=", 2)
    Next lIndex
    
    FileFromString "C:\David\Shared\Symbols2.TXT", astrSymbols.JoinFields(",")
    AddList "Done"
    
        
'    SearchInProjectFiles "tmrRealTime"
'    Dim Basket As New cStrategyBasket
'
'    Basket.Load "C:\David\Shared\A Test.SB" 'AddSlash(App.Path) & "Custom\A Test.SB"  '"C:\David\Shared\A Test.SB"
'    Basket.SaveDb
#End If
        
#If 0 Then
    Dim strGenesisSymbol As String
    Dim strCqgSymbol As String
    
    strCqgSymbol = "F.US.GLEG13"
    strGenesisSymbol = g.RjoCqg.GenesisSymbol(strCqgSymbol, "")
    AddList strCqgSymbol & " ==> '" & strGenesisSymbol & "'"
    AddList strCqgSymbol & " ==> '" & g.RjoCqg.GenesisSymbol(strCqgSymbol, "", , True) & "'"
    AddList strGenesisSymbol & " ==> '" & g.RjoCqg.BrokerSymbol(strGenesisSymbol)
    
    strCqgSymbol = "F.US.GLES1G13"
    strGenesisSymbol = g.RjoCqg.GenesisSymbol(strCqgSymbol, "")
    AddList strCqgSymbol & " ==> '" & strGenesisSymbol & "'"
    AddList strCqgSymbol & " ==> '" & g.RjoCqg.GenesisSymbol(strCqgSymbol, "", , True) & "'"
    AddList strGenesisSymbol & " ==> '" & g.RjoCqg.BrokerSymbol(strGenesisSymbol)
    
    strCqgSymbol = "C.US.GLEH1313200"
    strGenesisSymbol = g.RjoCqg.GenesisSymbol(strCqgSymbol, "")
    AddList strCqgSymbol & " ==> '" & strGenesisSymbol & "'"
    AddList strCqgSymbol & " ==> '" & g.RjoCqg.GenesisSymbol(strCqgSymbol, "", , True) & "'"
    AddList strGenesisSymbol & " ==> '" & g.RjoCqg.BrokerSymbol(strGenesisSymbol)
    
    strCqgSymbol = "P.US.GLEH1313400"
    strGenesisSymbol = g.RjoCqg.GenesisSymbol(strCqgSymbol, "")
    AddList strCqgSymbol & " ==> '" & strGenesisSymbol & "'"
    AddList strCqgSymbol & " ==> '" & g.RjoCqg.GenesisSymbol(strCqgSymbol, "", , True) & "'"
    AddList strGenesisSymbol & " ==> '" & g.RjoCqg.BrokerSymbol(strGenesisSymbol)
#End If
    
#If 0 Then
    Dim dNow As Double                  ' Current time
    Dim lSessionDate As Long            ' Session date for the current time
    Dim dOnCloseTimeExch As Double      ' "On-close" time in exchange time
    Dim dStartTimeExch As Double        ' Session start time in exchange time
    Dim bBetween As Boolean             ' Are we between the "on-close" time and the next start time?
    Dim Bars As cGdBars                 ' Bars object
    Dim lDate As Long
    Dim adTimes As cGdArray
    Dim lIndex As Long
    
    Set Bars = New cGdBars
    SetBarProperties Bars, "ES-201303"
    Set adTimes = New cGdArray
    adTimes.Create eGDARRAY_Doubles, 7
    
    adTimes(0) = (7# / 24#)
    adTimes(1) = (17# / 24#) + (12# / 1440#)    ' On-Close time
    adTimes(2) = (17# / 24#) + (25# / 1440#)
    adTimes(3) = (17# / 24#) + (30# / 1440#)    ' Crossover time
    adTimes(4) = (17# / 24#) + (35# / 1440#)
    adTimes(5) = (18# / 24#)                    ' Opening time
    adTimes(6) = (20# / 24#)
    
    m.dOnCloseTimeExch = (17# / 24#) + (12# / 1440#)
    
    For lDate = JulFromLong(20130106) To JulFromLong(20130112)
        For lIndex = 0 To 5
    
            ' Get current time in exchange time...
            dNow = Val(lDate) + adTimes(lIndex)
            
            ' Get current session date...
            lSessionDate = Bars.SessionDateForTime(dNow, True)
            If lSessionDate <= 0 Then
                ' If not a valid normal trading time, then we must be between...
                bBetween = True
                AddList "Between set to " & Str(bBetween) & " ( Current Exchange Time = " & DateFormat(dNow, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ", On Close Time = " & DateFormat(m.dOnCloseTimeExch, NO_DATE, HH_MM_SS, AMPM_UPPER) & ", Start Time = " & DateFormat(Bars.Prop(eBARS_StartTime) / 1440#, NO_DATE, HH_MM_SS, AMPM_UPPER) & " ) -- Not valid trading time"
            Else
                ' Find the on-close time for the current trading session...
                dOnCloseTimeExch = Val(lSessionDate) + m.dOnCloseTimeExch
                
                ' Find the start time for the current trading session ...
                dStartTimeExch = lSessionDate + (Bars.Prop(eBARS_StartTime) / 1440#)
                If Bars.Prop(eBARS_StartTime) > Bars.Prop(eBARS_EndTime) Then
                    dStartTimeExch = dStartTimeExch - 1
                End If
                
                bBetween = ((dNow < dStartTimeExch) Or (dNow >= dOnCloseTimeExch))
                AddList "Between set to " & Str(bBetween) & " ( Current Exchange Time = " & DateFormat(dNow, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ", On Close Time = " & DateFormat(dOnCloseTimeExch, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & ", Start Time = " & DateFormat(dStartTimeExch, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & " )"
            End If
        
        Next lIndex
    Next lDate
#End If
        
#If 0 Then
    Dim Bars As cGdBars
    Dim dNow As Double
    Dim lSessionDate As Long
    
    Set Bars = New cGdBars
    SetBarProperties Bars, "ES-201303"
    
    dNow = 41277# + (17# / 24#) + (30# / 1440#)
    AddList "Now = " & DateFormat(dNow, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER)
    
    lSessionDate = Bars.SessionDateForTime(dNow, True)
    If lSessionDate <= 0 Then
        AddList "Session Date = " & Str(lSessionDate)
    Else
        AddList "Session Date = " & DateFormat(lSessionDate, MM_DD_YYYY)
    End If
#End If
        
#If 0 Then
    Dim astrFile As cGdArray
    Dim astrLine As cGdArray
    Dim astrOutput As cGdArray
    Dim lIndex As Long
    
    Set astrFile = New cGdArray
    Set astrOutput = New cGdArray
    astrOutput.Create eGDARRAY_Strings
    
    If astrFile.FromFile(AddSlash(App.Path) & "Info\SymbolMap.CSV") Then
        For lIndex = 0 To astrFile.Size - 1
            Set astrLine = New cGdArray
            astrLine.SplitFields astrFile(lIndex), ","
            
            If (Len(astrLine(1)) > 0) And (Len(astrLine(2)) > 0) Then
                astrOutput.Add "Call CpyRolls " & astrLine(1) & " " & astrLine(2)
            End If
        Next lIndex
    End If
    
    astrOutput.Sort
    astrOutput.ToFile AddSlash(App.Path) & "CpySynth.BAT"
#End If
        
#If 0 Then
    Dim Order1 As cPtOrder
    Dim Order2 As cPtOrder
    Dim Bars As cGdBars
    Dim strSymbol As String
    Dim dEntryPrice As Double
    Dim lAccountID As Long
    Dim strAccountNumber As String
    Dim lTriggerOrderID As Long
    
    strSymbol = "ES-201212"
    dEntryPrice = 1422
    lTriggerOrderID = 69
    
    strAccountNumber = "16800110" ' "SIM0001"  ' "16800110" -- PSTradeNav
    lAccountID = g.Broker.AccountIDForNumber(strAccountNumber)
    
    Set Bars = New cGdBars
    SetBarProperties Bars, strSymbol
    
    Set Order1 = New cPtOrder
    With Order1
        .AccountID = lAccountID
        .GenesisOrderID = NextGenesisOrderID(strAccountNumber, .Broker)
        .AutoTradeItemID = 0
        .Buy = False
        .Quantity = 1
        .SymbolOrSymbolID = strSymbol
        .Expiration = -1&
        
        .OrderType = eTT_OrderType_Limit
        .LimitPrice = dEntryPrice + (Bars.TickMove * 5)
        .StopPrice = 0
    End With
    
    Set Order2 = New cPtOrder
    With Order2
        .AccountID = lAccountID
        .GenesisOrderID = NextGenesisOrderID(strAccountNumber, .Broker)
        .AutoTradeItemID = 0
        .Buy = False
        .Quantity = 1
        .SymbolOrSymbolID = strSymbol
        .Expiration = -1&
        
        .OrderType = eTT_OrderType_Stop
        .LimitPrice = 0
        .StopPrice = dEntryPrice - (Bars.TickMove * 5)
    End With
    
    g.Broker.SubmitOrdersAsOco Order1, Order2, lTriggerOrderID, False
#End If
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTest2.Command1_Click"

End Sub

Private Sub Command2_Click()
On Error GoTo ErrSection:
        
    Dim astrSymbols As cGdArray         ' Array of symbols to process
    Dim lIndex As Long                  ' Index into a for loop
    Dim strFileBase As String           ' Base for the files
    
    Set astrSymbols = New cGdArray
    If astrSymbols.FromFile(AddSlash(App.Path) & "IbForex.TXT") Then
        For lIndex = 0 To astrSymbols.Size - 1
            AddList astrSymbols(lIndex)
            
            strFileBase = AddSlash(App.Path) & "IbFx\" & Replace(astrSymbols(lIndex), "$", "")
            
            SecondBarFileToAsc strFileBase & ".BAR", strFileBase & ".ASC"
        Next lIndex
    End If
    
    AddList "Done"
        
#If 0 Then
    Dim strFileName As String
    Dim strLot As String
    Dim turnkeyMessage As cBrokerMessage
    
    If Not g.Turnkey Is Nothing Then
        strFileName = InfBox("What is the Feedyard Lot File?", "?", , "Feedyard Lot File", , , , , , "string")
        If Len(strFileName) > 0 Then
            strLot = FileToString(strFileName)
            If Len(strLot) > 0 Then
                Set turnkeyMessage = New cBrokerMessage
                turnkeyMessage.FromString strLot
                
                turnkeyMessage.Remove "FeedYardLotID"
                turnkeyMessage("FeedYardName") = frmTurnkey.cboFeedYards.Text
                
                g.Turnkey.AddLots turnkeyMessage.ToString
            End If
        End If
    End If
#End If
        
#If 0 Then
    Dim astrFile As cGdArray
    Dim astrLine As cGdArray
    Dim astrOutput As cGdArray
    Dim lIndex As Long
    Dim Bars As cGdBars
    Dim astrOutputLine As cGdArray
    Dim bDiff As Boolean
    Dim dStart55 As Double
    Dim dStart56 As Double
    Dim dStart57 As Double
    
    Set astrFile = New cGdArray
    Set astrOutput = New cGdArray
    astrOutput.Create eGDARRAY_Strings
    
    If astrFile.FromFile(AddSlash(App.Path) & "Info\SymbolMap.CSV") Then
        For lIndex = 0 To astrFile.Size - 1
            Set astrLine = New cGdArray
            astrLine.SplitFields astrFile(lIndex), ","
            
            If (Len(astrLine(1)) > 0) And (Len(astrLine(2)) > 0) Then
                Set astrOutputLine = New cGdArray
                astrOutputLine.Create eGDARRAY_Strings, 8
                
                AddList astrLine(1) & "/" & astrLine(2)
                astrOutputLine(0) = astrLine(1)
                astrOutputLine(4) = astrLine(2)
                
                bDiff = False
                Set Bars = New cGdBars
                
                dStart55 = 0
                dStart56 = 0
                dStart57 = 0
                
                If DM_GetBars(Bars, astrLine(1) & "-055") Then
                    astrOutputLine(1) = DateFormat(Bars(eBARS_DateTime, 0), MM_DD_YYYY)
                    dStart55 = Bars(eBARS_DateTime, 0)
                End If
                
                If DM_GetBars(Bars, astrLine(1) & "-056") Then
                    astrOutputLine(2) = DateFormat(Bars(eBARS_DateTime, 0), MM_DD_YYYY)
                    dStart56 = Bars(eBARS_DateTime, 0)
                End If
                
                If DM_GetBars(Bars, astrLine(1) & "-057") Then
                    astrOutputLine(3) = DateFormat(Bars(eBARS_DateTime, 0), MM_DD_YYYY)
                    dStart57 = Bars(eBARS_DateTime, 0)
                End If
            
                If DM_GetBars(Bars, astrLine(2) & "-055") Then
                    astrOutputLine(5) = DateFormat(Bars(eBARS_DateTime, 0), MM_DD_YYYY)
                End If
                If Bars(eBARS_DateTime, 0) < dStart55 Then
                    astrOutputLine(5) = astrOutputLine(5) & "*"
                    bDiff = True
                ElseIf Bars(eBARS_DateTime, 0) > dStart55 Then
                    astrOutputLine(5) = astrOutputLine(5) & "**"
                    bDiff = True
                End If
                
                If DM_GetBars(Bars, astrLine(2) & "-056") Then
                    astrOutputLine(6) = DateFormat(Bars(eBARS_DateTime, 0), MM_DD_YYYY)
                End If
                If Bars(eBARS_DateTime, 0) < dStart56 Then
                    astrOutputLine(6) = astrOutputLine(6) & "*"
                    bDiff = True
                ElseIf Bars(eBARS_DateTime, 0) > dStart56 Then
                    astrOutputLine(6) = astrOutputLine(6) & "**"
                    bDiff = True
                End If
                
                If DM_GetBars(Bars, astrLine(2) & "-057") Then
                    astrOutputLine(7) = DateFormat(Bars(eBARS_DateTime, 0), MM_DD_YYYY)
                End If
                If Bars(eBARS_DateTime, 0) < dStart57 Then
                    astrOutputLine(7) = astrOutputLine(7) & "*"
                    bDiff = True
                ElseIf Bars(eBARS_DateTime, 0) > dStart57 Then
                    astrOutputLine(7) = astrOutputLine(7) & "**"
                    bDiff = True
                End If
                
                If bDiff = True Then
                    AddList astrOutputLine.JoinFields(";")
                End If
                
                astrOutput.Add astrOutputLine.JoinFields(vbTab)
            End If
        Next lIndex
    End If
    
    astrOutput.Sort
    astrOutput.Add "ELEC" & vbTab & "55" & vbTab & "56" & vbTab & "57" & vbTab & "SYNTH" & vbTab & "55" & vbTab & "56" & vbTab & "57", 0
    
    astrOutput.ToFile AddSlash(App.Path) & "Synth.TXT"
    AddList "Done"
#End If
        
'    Dim astrOptions As cGdArray         ' Contingency options for the order
'    Dim Order As cPtOrder               ' Order object
'    Dim strSymbol As String
'    Dim dEntryPrice As Double
'    Dim lAccountID As Long
'    Dim strAccountNumber As String
'
'    strSymbol = "ES-201212"
'    dEntryPrice = 1420
'
'    strAccountNumber = "16800110" ' "SIM0001"  ' "16800110" -- PSTradeNav
'    lAccountID = g.Broker.AccountIDForNumber(strAccountNumber)
'
'    Set astrOptions = New cGdArray
'    astrOptions.Create eGDARRAY_Strings, 10
'
'    astrOptions(0) = "-1"
'    astrOptions(1) = "0"
'    astrOptions(2) = "62.50"
'    astrOptions(3) = "-1"
'    astrOptions(4) = "1.25"
'
'    astrOptions(5) = "-1"
'    astrOptions(6) = "0"
'    astrOptions(7) = "62.50"
'    astrOptions(8) = "-1"
'    astrOptions(9) = "1.25"
'
'    Set Order = New cPtOrder
'    With Order
'        .AccountID = lAccountID
'        .GenesisOrderID = NextGenesisOrderID(strAccountNumber, .Broker)
'        .AutoTradeItemID = 0
'        .Buy = True
'        .Quantity = 1
'        .SymbolOrSymbolID = strSymbol
'        .Expiration = -1&
'
'        .OrderType = eTT_OrderType_Limit
'        .LimitPrice = dEntryPrice
'        .StopPrice = 0
'
'        .ContingencyOptions = astrOptions.JoinFields(",")
'    End With
'
'    SubmitOrder Order
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTest2.Command2_Click"

End Sub

Private Sub Command3_Click()
On Error GoTo ErrSection:
        
#If 0 Then
    Dim strYard As String
    Dim strCode As String
    
    If Not g.Turnkey Is Nothing Then
        strYard = InfBox("What is the Feedyard Name?", "?", , "Feedyard Name", , , , , , "string")
        If Len(strYard) > 0 Then
            strCode = InfBox("What is the Feedyard Code?", "?", , "Feedyard Code", , , , , , "string")
            If Len(strCode) > 0 Then
                g.Turnkey.AddFeedyards "Code=" & strCode & vbTab & "Name=" & strYard
            End If
        End If
    End If
#End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.Command3_Click"
    
End Sub

Private Sub Command4_Click()
On Error GoTo ErrSection:
    
#If 0 Then
    If Not g.Turnkey Is Nothing Then
        g.Turnkey.GetLotColumnCategories
    End If
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.Command4_Click"
    
End Sub

Private Sub Command5_Click()
On Error GoTo ErrSection:
    
#If 0 Then
    Dim strLot As String
        
    If Not g.Turnkey Is Nothing Then
        strLot = InfBox("What is the Lot ID?", "?", , "Lot ID", , , , , , "string")
        If Len(strLot) > 0 Then
            g.Turnkey.GetLotContentDetailsForLot strLot
        End If
    End If
#End If
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.Command5_Click"

End Sub

Private Sub Command6_Click()
On Error GoTo ErrSection:

#If 0 Then
    Dim strYard As String
        
    If Not g.Turnkey Is Nothing Then
        strYard = InfBox("What is the Feedyard ID?", "?", , "Feedyard ID", , , , , , "string")
        If Len(strYard) > 0 Then
            g.Turnkey.GetLotContentDetailsForYard strYard
        End If
    End If
#End If
    
    'g.RealTime.bDisableForTesting = Not g.RealTime.bDisableForTesting

ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.Command6_Click"

End Sub

Private Sub Command7_Click()
On Error GoTo ErrSection:

    AnalyzeFuturesTable

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.Command7_Click"
    
End Sub

Private Sub Form_Activate()

'    tmrTest1.Enabled = True
'    Sleep 0.5
'    tmrTest2.Enabled = True

End Sub

Private Sub Form_Load()

    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
'    Command2.Caption = "Add Lot"
'    Command3.Caption = "Add Yard"
'    Command4.Caption = "Categories"
'    Command5.Caption = "Details/Lot"
'    Command6.Caption = "Details/Yard"
    Command7.Caption = "Fut Tbl"
    
    tmrTest1.Enabled = False
    tmrTest1.Interval = 1000
    
    tmrTest2.Enabled = False
    tmrTest2.Interval = 1000
    
    Dim Bars As cGdBars
    
    Set Bars = New cGdBars
    SetBarProperties Bars, "ES-201212"
    
    pePrice.Init Bars
    
    DoClearColors = True
    
    m.bIgnoreAppLoaded = False
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    lst.Height = Me.ScaleHeight - lst.Top * 2
    lst.Width = Me.ScaleWidth - lst.Left - lst.Top

End Sub

Private Sub pePrice_Changed()

    frmTest2.AddList "New Value = " & Str(pePrice.Price)

End Sub

Private Sub tmrDownload_Timer()
On Error GoTo ErrSection:

    Static nPrevStatus As eGDDownloadStatus
    Dim nStatus As eGDDownloadStatus
    Dim strError As String
    
    If Not m.Downloader Is Nothing Then
        nStatus = m.Downloader.Status(strError)
        
        If (nPrevStatus = 0) Or (nStatus <> nPrevStatus) Then
            Select Case nStatus
                Case eGDDownloadStatus_Nothing
                    AddList "Nothing"
                Case eGDDownloadStatus_Downloading
                    AddList "Downloading"
                Case eGDDownloadStatus_Done
                    AddList "Done"
                Case eGDDownloadStatus_Aborted
                    AddList "Aborted"
                Case eGDDownloadStatus_Error
                    AddList "Error: '" & strError & "'"
            End Select
            
            If (nStatus <> eGDDownloadStatus_Downloading) And (nStatus <> eGDDownloadStatus_Nothing) Then
                tmrDownload.Enabled = False
            End If
        End If
        
        nPrevStatus = nStatus
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.tmrDownload_Timer"
    
End Sub

Private Sub tmrTest1_Timer()

    AddList "tmrTest1_Timer: Showing Modeless InfBox"
    InfBox "This is a modeless dialog", , "+-OK", "Modeless Dialog", True
    AddList "tmrTest1_Timer: Done showing Modeless InfBox"

End Sub

Private Sub tmrTest2_Timer()

    AddList "tmrTest2_Timer: Showing Modal InfBox"
#If 1 Then
    InfBox "This is a modal dialog", , "+OK|-Cancel", "Modal Dialog"
#Else
    If InfBox("This is a modal dialog", , "+OK|-Stop", "Modal Dialog") = "S" Then
        tmrTest2.Enabled = False
        tmrTest1.Enabled = False
    End If
#End If
    AddList "tmrTest2_Timer: Done showing Modal InfBox"

End Sub

Public Sub AddList(ByVal strMsg$, Optional ByVal bBenchMark As Boolean = False)
    
    Dim dTicks#, i&
    Static dPrevTicks#
    
    If g.bUnloading Then Exit Sub
    If (Not IsIDE) And (Not Me.Visible) Then Exit Sub
    
    If bBenchMark Then
        dTicks = gdTickCount
        If Len(strMsg) > 0 Then
            strMsg = strMsg & ":  " & Format((dTicks - dPrevTicks) / 1000, "0.000") & " seconds"
        End If
    End If
    
    If Len(strMsg) > 0 Then
        With lst
            If .ListCount > 2000 Then
                .AddItem "*** REMOVING LINES FROM LISTBOX ***"
                .ListIndex = .ListCount - 1
'RH commented out                 .Refresh
                For i = 1000 To 0 Step -1
                    .RemoveItem i
                Next
            End If
            .AddItem strMsg
            .ListIndex = .ListCount - 1
'RH commented out             .Refresh
        End With
    End If
    dPrevTicks = gdTickCount
    
    DebugLog strMsg

End Sub

Private Sub lst_DblClick()

    lst.Clear
    'RH commented out lst.Refresh

End Sub

Public Sub GenerateReport(ByVal vArgs As Variant)

    With frmPrintPreview.vp
        .StartDoc
        
        .LineSpacing = 100
        .HdrFontName = "Times New Roman"
        .HdrFontSize = 14
        .Header = "|Chart Navigator" & vbCrLf & "Genesis Financial Data Services - (800) 808-DATA - www.gfds.com"
        .Footer = "|Page: %d|"
        
        .DrawPicture LoadPicture(AddSlash(App.Path) & "Temp.BMP"), .MarginLeft, .MarginTop, , , vppaZoom

        .EndDoc
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AnalyzeFuturesTable
'' Description: Analyze the futures table to determine available CSI numbers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AnalyzeFuturesTable()
On Error GoTo ErrSection:

    Dim astrFile As New cGdArray        ' Futures Table file read into an array
    Dim astrTable As New cGdArray       ' Table
    Dim lIndex As Long                  ' Index into a for loop
    Dim strCsiNum As String             ' CSI Number
    Dim lCsiNum As Long                 ' CSI Number
    Dim strFileSymbol As String         ' Symbol from the file
    Dim strTblSymbol As String          ' Symbol from the table
    Dim astrOutput As New cGdArray      ' Output array
    Dim strToOutput As String           ' String to output
    
    astrTable.Create eGDARRAY_Strings, 1000
    
    If astrFile.FromFile("K:\Common\Futures.TBL") Then
        For lIndex = 0 To astrFile.Size - 1
            If Left(astrFile(lIndex), 1) <> ";" Then
                strCsiNum = Trim(Mid(astrFile(lIndex), 5, 3))
                lCsiNum = CLng(Val(strCsiNum))
                
                If Len(astrTable(lCsiNum)) > 0 Then
                    If lCsiNum <> 0 Then
                        strTblSymbol = Trim(Left(astrTable(lCsiNum), 3))
                        strFileSymbol = Trim(Left(astrFile(lIndex), 3))
                    
                        strToOutput = "Conflict with " & strCsiNum & " (" & strTblSymbol & ", " & strFileSymbol & ")"
                        astrOutput.Add strToOutput
                        frmTest2.AddList strToOutput
                    End If
                Else
                    astrTable(lCsiNum) = astrFile(lIndex)
                End If
            End If
        Next lIndex
        
        For lIndex = 0 To astrTable.Size - 1
            If Len(astrTable(lIndex)) = 0 Then
                strToOutput = "Available: " & Str(lIndex)
                astrOutput.Add strToOutput
                frmTest2.AddList strToOutput
            End If
        Next lIndex
        
        astrOutput.ToFile AddSlash(App.Path) & "FutTbl.TXT"
    End If
    
    frmTest2.AddList "Done"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.AnalyzeFuturesTable"
    
End Sub

#If 0 Then
Private Function PropertyNameForSymbol(ByVal strSymbol As String) As String

    Dim strReturn As String
    Dim strBrokerCode As String

    strReturn = ""
    If IsForex(strSymbol) Then
        If InStr(strSymbol, "@") <> 0 Then
            strBrokerCode = Parse(strSymbol, "@", 2)
            strBrokerCode = UCase(Left(strBrokerCode, 1)) & LCase(Mid(strBrokerCode, 2))
            strReturn = strBrokerCode & "Fx"
        End If
    End If
    
    PropertyNameForSymbol = strReturn

End Function

Private Function IsBetween(Bars As cGdBars, ByVal dExchOnCloseTime As Double) As Boolean

    Dim bReturn As Boolean
    Dim dNow As Double
    Dim lSessionDate As Long
    Dim dStartTimeExch As Double
    Dim dEndTimeExch As Double
    
    ' Get current time in exchange time...
    dNow = CurrentTime(Bars.Prop(eBARS_ExchangeTimeZoneInf))
    
    ' Get current session date...
    lSessionDate = Bars.SessionDateForTime(dNow, True)
    If lSessionDate <= 0 Then
        ' If not a valid normal trading time, then we must be between...
        bReturn = True
    Else
        ' Find the on close time for the current trading session...
        dExchOnCloseTime = Val(lSessionDate) + dExchOnCloseTime
        
        ' Find the start time for the current trading session ...
        dStartTimeExch = lSessionDate + (Bars.Prop(eBARS_StartTime) / 1440#)
        If Bars.Prop(eBARS_StartTime) > Bars.Prop(eBARS_EndTime) Then
            dStartTimeExch = dStartTimeExch - 1
        End If
        
        bReturn = (dNow < dStartTimeExch Or dNow >= dExchOnCloseTime)
        
        
        ' Assume on-close time is between midnight exchange and the normal end time...
        
#If 0 Then
        ' Find the start time for the next trading session...
        Do
            lSessionDate = lSessionDate + 1
        While Not IsWeekday(lSessionDate)
        
        dStartTimeExch = lSessionDate + (Bars.Prop(eBARS_StartTime) / 1440#)
        
        If Bars.Prop(eBARS_StartTime) > Bars.Prop(eBARS_EndTime) Then
            dStartTimeExch = dStartTimeExch - 1
        End If
        
        bReturn = (dNow >= dExchOnCloseTime) And (dNow < dStartTimeExch)
#End If
    End If
    
    IsBetween = bReturn

End Function

Private Function ProjectFiles() As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of files to return
    Dim astrProject As cGdArray         ' Project file read into an array
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    Set astrProject = New cGdArray
    
    If astrProject.FromFile(AddSlash(App.Path) & "NavSuite.VBP") Then
        For lIndex = 0 To astrProject.Size - 1
            If UCase(Left(astrProject(lIndex), 5)) = "FORM=" Then
                astrReturn.Add AddSlash(App.Path) & Trim(Parse(astrProject(lIndex), "=", 2))
            ElseIf UCase(Left(astrProject(lIndex), 6)) = "CLASS=" Then
                astrReturn.Add AddSlash(App.Path) & Trim(Parse(astrProject(lIndex), ";", 2))
            ElseIf UCase(Left(astrProject(lIndex), 7)) = "MODULE=" Then
                astrReturn.Add AddSlash(App.Path) & Trim(Parse(astrProject(lIndex), ";", 2))
            End If
        Next lIndex
    End If
    
    Set ProjectFiles = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTest2.ProjectFiles"
    
End Function

Private Sub SearchInProjectFiles(ByVal strSearchString As String, Optional ByVal strDumpFile As String = "")
On Error GoTo ErrSection:

    Dim astrProjectFiles As cGdArray    ' List of files in the project
    Dim lFile As Long                   ' Index into a for loop
    Dim astrFile As cGdArray            ' File read into an array
    Dim lLine As Long                   ' Index into a for loop
    Dim astrOutput As cGdArray          ' Output array
    Dim strOutput As String             ' String to add to the output
    
    Set astrProjectFiles = ProjectFiles
    Set astrOutput = New cGdArray
    
    For lFile = 0 To astrProjectFiles.Size - 1
        Set astrFile = New cGdArray
        If astrFile.FromFile(astrProjectFiles(lFile)) Then
            For lLine = 0 To astrFile.Size - 1
                If InStr(astrFile(lLine), strSearchString) <> 0 Then
                    strOutput = astrProjectFiles(lFile) & " (Line " & Str(lLine + 1) & "): " & astrFile(lLine)
                    
                    AddList strOutput
                    astrOutput.Add strOutput
                End If
            Next lLine
        End If
    Next lFile
    
    If Len(strDumpFile) > 0 Then
        astrOutput.ToFile strDumpFile
    End If
    
    AddList "Done"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.SearchInProjectFiles"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidTradingTime
'' Description: Is this a valid trading time for the symbol?
'' Inputs:      None
'' Returns:     True if valid trading time, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidTradingTime(Optional dTime As Double = kNullData) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strTradeSymbol As String        ' Trade symbol
    Dim dCurrentTime As Double          ' Current time
    
    bReturn = True
    strTradeSymbol = TradeSymbol
    
    If (IsForex(strTradeSymbol) = True) And (InStr(strTradeSymbol, "@") <> 0) Then
        If dTime = kNullData Then
            dCurrentTime = CurrentTime("NY", strTradeSymbol, True)
        Else
            dCurrentTime = dTime
        End If
        
        bReturn = mDataNav.ValidForexTradingTime(strTradeSymbol, dCurrentTime)
    Else
        If dTime = kNullData Then
            dCurrentTime = CurrentTime(m.Bars.Prop(eBARS_ExchangeTimeZoneInf), TradeSymbol, True)
        Else
            dCurrentTime = dTime
        End If
        
        Select Case Weekday(dCurrentTime)
            Case vbFriday
                bReturn = (dCurrentTime <= (Val(Int(dCurrentTime)) + (m.Bars.Prop(eBARS_StartTime) / 1440#)))
            Case vbSunday
                bReturn = (dCurrentTime >= (Val(Int(dCurrentTime)) + (m.Bars.Prop(eBARS_EndTime) / 1440#)))
            Case vbSaturday
                bReturn = False
            Case Else
                bReturn = True
        End Select
    End If
        
    ValidTradingTime = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cActiveTsOrderGroup.ValidTradingTime"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CombinedRollFiles
'' Description: Build combined roll files for symbols that have a pit/elec/combined
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CombinedRollFiles()
On Error GoTo ErrSection:

    Dim astrSymbolMap As cGdArray       ' Symbol map
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lIndex3 As Long                 ' Index into a for loop
    Dim astrComponents As cGdArray      ' Components of the symbol map
    Dim astrPit As cGdArray             ' Array of pit rolls
    Dim astrCombined As cGdArray        ' Array of combined rolls
    Dim astrElectronic As cGdArray      ' Array of electronic rolls
    Dim lFirstCombined As Long          ' First valid combined contract
    Dim lFirstElectronic As Long        ' First valid electronic contract
    Dim lFirstPit As Long               ' First valid pit contract
    Dim strRollFile As String           ' Roll file
    Dim lContract As Long               ' Current contract
    Dim astrOutput As cGdArray          ' Output array
    Dim strRollPath As String           ' Base path for the rolls
    
    Set astrSymbolMap = New cGdArray
    astrSymbolMap.FromFile AddSlash(App.Path) & "Info\SymbolMap.CSV"
    
    strRollPath = "C:\Rolls\"
    
    For lIndex = 0 To astrSymbolMap.Size - 1
        ' Pit, Elec, Synth, Comb
        Set astrComponents = New cGdArray
        astrComponents.SplitFields astrSymbolMap(lIndex), ","
        
        ' Has electronic AND either pit OR combined...
        If (Len(astrComponents(1)) > 0) And ((Len(astrComponents(0)) > 0) Or (Len(astrComponents(3)) > 0)) Then
            For lIndex2 = 55 To 57
                strRollFile = strRollPath & astrComponents(1) & "-99" & Str(lIndex2) & ".ROL"
                
                AddList "Loading " & strRollFile & "..."
                Set astrElectronic = New cGdArray
                astrElectronic.FromFile strRollFile
                lFirstElectronic = CLng(Val(Parse(astrElectronic(1), " ", 2)))
                
                If Len(astrComponents(0)) > 0 Then
                    strRollFile = strRollPath & astrComponents(0) & "-99" & Str(lIndex2) & ".ROL"
                    
                    AddList "Loading " & strRollFile & "..."
                    Set astrPit = New cGdArray
                    astrPit.FromFile strRollFile
                    lFirstPit = CLng(Val(Parse(astrPit(1), " ", 2)))
                Else
                    Set astrPit = Nothing
                    lFirstPit = 99999
                End If
                
                If Len(astrComponents(3)) > 0 Then
                    strRollFile = strRollPath & astrComponents(3) & "-99" & Str(lIndex2) & ".ROL"
                    
                    AddList "Loading " & strRollFile & "..."
                    Set astrCombined = New cGdArray
                    astrCombined.FromFile strRollFile
                    lFirstCombined = CLng(Val(Parse(astrCombined(1), " ", 2)))
                Else
                    Set astrCombined = Nothing
                    lFirstCombined = 99999
                End If
                
                If (lFirstCombined < lFirstElectronic) Or (lFirstPit < lFirstElectronic) Then
                    Set astrOutput = New cGdArray
                    astrOutput.CopyFromHandle astrElectronic.ArrayHandle
                    astrOutput.Remove 0
                    
                    If lFirstCombined < lFirstElectronic Then
                        For lIndex3 = astrCombined.Size - 1 To 1 Step -1
                            lContract = CLng(Val(Parse(astrCombined(lIndex3), " ", 2)))
                            If lContract < lFirstElectronic Then
                                AddList "     Prepending " & Str(lContract) & " from combined..."
                                astrOutput.Add astrCombined(lIndex3), 0
                            End If
                        Next lIndex3
                        
                        lFirstElectronic = lFirstCombined
                    End If
                    
                    If lFirstPit < lFirstElectronic Then
                        For lIndex3 = astrPit.Size - 1 To 0 Step -1
                            lContract = CLng(Val(Parse(astrPit(lIndex3), " ", 2)))
                            If lContract < lFirstElectronic Then
                                AddList "     Prepending " & Str(lContract) & " from pit..."
                                astrOutput.Add astrPit(lIndex3), 0
                            End If
                        Next lIndex3
                    Else
                        astrOutput.Add astrCombined(0), 0
                    End If
                    
                    strRollFile = AddSlash(App.Path) & "Rolls\" & astrComponents(1) & "-99" & Str(lIndex2) & ".ROL"
                    AddList "Writing " & strRollFile & "..."
                    astrOutput.ToFile strRollFile
                End If
            Next lIndex2
        End If
    Next lIndex
    
    AddList "Done"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.CombinedRollFiles"
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadMjkToGen
'' Description: Load the Mjk to Genesis conversion tree
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadMjkToGen()
On Error GoTo ErrSection:

    Dim astrFuturesTable As cGdArray    ' Futures table
    Dim lIndex As Long                  ' Index into a for loop
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim strMJKSymbol As String          ' MJK symbol
    Dim iPos As Integer                 ' Position of the something in the string
    
    Set m.MjkToGen = New cGdTree
    
    Set astrFuturesTable = New cGdArray
    If astrFuturesTable.FromFile("K:\Common\Futures.TBL") Then
        For lIndex = 0 To astrFuturesTable.Size - 1
            If Left(astrFuturesTable(lIndex), 1) <> ";" Then
                iPos = InStr(astrFuturesTable(lIndex), "=")
                If iPos <> 0 Then
                    strGenesisSymbol = Trim(Mid(astrFuturesTable(lIndex), iPos + 1, 6))
                    
                    iPos = InStr(iPos, astrFuturesTable(lIndex), "m ")
                    If iPos = 0 Then
                        strMJKSymbol = strGenesisSymbol
                    Else
                        strMJKSymbol = Trim(Mid(astrFuturesTable(lIndex), iPos + 2, 8))
                    End If
                    
                    m.MjkToGen.Add strGenesisSymbol, strMJKSymbol
                End If
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.LoadMjkToGen"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGenToMjk
'' Description: Load the Genesis to Mjk conversion tree
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGenToMjk()
On Error GoTo ErrSection:

    Dim astrFuturesTable As cGdArray    ' Futures table
    Dim lIndex As Long                  ' Index into a for loop
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim strMJKSymbol As String          ' MJK symbol
    Dim iPos As Integer                 ' Position of the something in the string
    
    Set m.GenToMjk = New cGdTree
    
    Set astrFuturesTable = New cGdArray
    If astrFuturesTable.FromFile("K:\Common\Futures.TBL") Then
        For lIndex = 0 To astrFuturesTable.Size - 1
            If Left(astrFuturesTable(lIndex), 1) <> ";" Then
                iPos = InStr(astrFuturesTable(lIndex), "=")
                If iPos <> 0 Then
                    strGenesisSymbol = Trim(Mid(astrFuturesTable(lIndex), iPos + 1, 6))
                    
                    iPos = InStr(iPos, astrFuturesTable(lIndex), "m ")
                    If iPos = 0 Then
                        strMJKSymbol = strGenesisSymbol
                    Else
                        strMJKSymbol = Trim(Mid(astrFuturesTable(lIndex), iPos + 2, 8))
                    End If
                    
                    m.GenToMjk.Add strMJKSymbol, strGenesisSymbol
                End If
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.LoadGenToMjk"
    
End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixTotVolAndOi
'' Description: Fix the total volume/total oi on last two days of every futures
''              contract after June 2009
'' Inputs:      Start Date, End Date, Base Symbols to Process
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixTotVolAndOi(ByVal lStartDate As Long, ByVal lEndDate As Long, Optional ByVal strBases As String = "")
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lIndex3 As Long                 ' Index into a for loop
    Dim strBaseSymbol As String         ' Base symbol
    Dim strSymbol As String             ' Symbol
    Dim lSymbolID As Long               ' Symbol ID for the selected symbol
    Dim astrMarkets As cGdArray         ' Array of futures markets
    Dim astrContracts As cGdArray       ' Array of futures contracts
    Dim lContract As Long               ' Contract for the future
    Dim Bars As cGdBars                 ' Bars object
    Dim Contracts As cGdTree            ' Collection of bars
    Dim alDates As cGdArray             ' Array of dates
    Dim alTotVol As cGdArray            ' Array of total volume
    Dim alTotOi As cGdArray             ' Array of total open interest
    Dim lBarNum As Long                 ' Bar number for the date
    Dim lCounter As Long                ' Counter variable
    Dim astrOutput As cGdArray          ' Output array
    Dim astrOutput2 As cGdArray         ' Output array
    Dim astrFields As cGdArray          ' Array of fields to output
    Dim astrZeroOi As cGdArray          ' Array of zero contract open interest
    
    Set astrMarkets = New cGdArray
    astrMarkets.Create eGDARRAY_Strings
    
    Set astrContracts = New cGdArray
    astrContracts.Create eGDARRAY_Strings
    
    Set alDates = New cGdArray
    alDates.Create eGDARRAY_Longs
    For lIndex = lStartDate To lEndDate
        alDates.Add lIndex
    Next lIndex
    
    Set alTotVol = New cGdArray
    alTotVol.Create eGDARRAY_Longs, alDates.Size
    
    Set alTotOi = New cGdArray
    alTotOi.Create eGDARRAY_Longs, alDates.Size
    
    Set astrOutput = New cGdArray
    astrOutput.Create eGDARRAY_Strings, alDates.Size
    
    Set astrOutput2 = New cGdArray
    astrOutput2.Create eGDARRAY_Strings
    
    Set astrFields = New cGdArray
    astrFields.Create eGDARRAY_Strings, 9
    
    Set astrZeroOi = New cGdArray
    astrZeroOi.Create eGDARRAY_Strings
    
    Set Contracts = New cGdTree
    
    If SU_GetMarkets(astrMarkets) Then
        astrMarkets.Sort
        
        ' Base;SymbolID;Description ( e.g. AC-0;77068;Ethanol CBOT (Pit) - Market; )
        For lIndex = 0 To astrMarkets.Size - 1
            AddList "Market = " & astrMarkets(lIndex)
            
            Contracts.Clear
            
            strBaseSymbol = Parse(Parse(astrMarkets(lIndex), ";", 1), "-", 1)
            lSymbolID = CLng(ValOfText(Parse(astrMarkets(lIndex), ";", 2)))
            
            'strBaseSymbol = "ES"
            'lSymbolID = GetSymbolID("ES-0")
            
            If (Len(strBases) = 0) Or (InStr(strBases, "," & strBaseSymbol & ",") <> 0) Then
                If SU_GetContracts(lSymbolID, astrContracts) Then
                    ' Symbol;Symbol ID ( e.g. AC-200506;77069 )
                    For lIndex2 = 0 To astrContracts.Size - 1
                        AddList "     Loading " & astrContracts(lIndex2)
                        
                        strSymbol = Parse(astrContracts(lIndex2), ";", 1)
                        lContract = CLng(Val(Parse(strSymbol, "-", 2)))
                        lSymbolID = CLng(Val(Parse(astrContracts(lIndex2), ";", 2)))
                        
                        Set Bars = New cGdBars
                        If DM_GetBars(Bars, lSymbolID, "Daily") Then
                            If Bars(eBARS_DateTime, Bars.Size - 1) >= lStartDate Then ' JulFromLong(20090630) Then
                                AddList vbTab & "Contract = " & astrContracts(lIndex2)
                                
                                Set Bars = New cGdBars
                                If DM_GetBars(Bars, strSymbol, "Daily") Then
                                    Contracts.Add Bars, Bars.Prop(eBARS_Symbol)
                                End If
                            End If
                        End If
                    Next lIndex2
                    
                    For lIndex2 = 0 To alDates.Size - 1
                        alTotVol(lIndex2) = 0
                        alTotOi(lIndex2) = 0
                        astrOutput(lIndex2) = ""
                    Next lIndex2
                    astrOutput2.Clear
                                      
                    For lIndex2 = 1 To Contracts.Count
                        Set Bars = Contracts(lIndex2)
                        AddList "     Walking through " & Bars.Prop(eBARS_Symbol) & " to calculate totals"
                        
                        For lIndex3 = 0 To Bars.Size - 1
                            If alDates.BinarySearch(Bars(eBARS_DateTime, lIndex3), lCounter) Then
                                alTotVol(lCounter) = alTotVol(lCounter) + Bars(eBARS_ContVol, lIndex3)
                                alTotOi(lCounter) = alTotOi(lCounter) + Bars(eBARS_ContOI, lIndex3)
                            End If
                        Next lIndex3
                    Next lIndex2
                    
                    For lIndex2 = 0 To alDates.Size - 1
                        astrOutput(lIndex2) = Str(alDates(lIndex2)) & vbTab & Str(JulToLong(alDates(lIndex2), -1)) & vbTab & Str(alTotVol(lIndex2)) & vbTab & Str(alTotOi(lIndex2))
                    Next lIndex2
                    astrOutput.ToFile AddSlash(App.Path) & "Totals\" & strBaseSymbol & ".TXT"
                    
                    For lIndex2 = 1 To Contracts.Count
                        Set Bars = Contracts(lIndex2)
                        AddList "     Walking through " & Bars.Prop(eBARS_Symbol) & " to fix data"
                        astrZeroOi.Add Bars.Prop(eBARS_Symbol)
                        
                        astrOutput2.Add "#" & Bars.Prop(eBARS_Symbol)
                        For lIndex3 = 0 To Bars.Size - 1
                            If alDates.BinarySearch(Bars(eBARS_DateTime, lIndex3), lCounter) Then
                                If (alTotVol(lCounter) <> Bars(eBARS_Vol, lIndex3)) Or (alTotOi(lCounter) <> Bars(eBARS_OI, lIndex3)) Then
                                    astrFields.Clear
                                    
                                    astrFields(0) = Format(Bars(eBARS_DateTime, lIndex3), "YYYYMMDD")
                                    astrFields(1) = Str(Bars(eBARS_Open, lIndex3))
                                    astrFields(2) = Str(Bars(eBARS_High, lIndex3))
                                    astrFields(3) = Str(Bars(eBARS_Low, lIndex3))
                                    astrFields(4) = Str(Bars(eBARS_Close, lIndex3))
                                    astrFields(5) = Str(Bars(eBARS_ContVol, lIndex3))
                                    astrFields(6) = Str(Bars(eBARS_ContOI, lIndex3))
                                    astrFields(7) = Str(alTotVol(lCounter))
                                    astrFields(8) = Str(alTotOi(lCounter))
                                    
                                    astrOutput2.Add astrFields.JoinFields(" ")
                                    
                                    If Bars(eBARS_ContOI, lIndex3) = 0 Then
                                        astrZeroOi.Add astrFields.JoinFields(" ")
                                    End If
                                End If
                            End If
                        Next lIndex3
                    Next lIndex2
                    
                    astrOutput2.ToFile AddSlash(App.Path) & "Totals\" & strBaseSymbol & ".DAT"
                End If
            End If
            
            DoEvents
        Next lIndex
        
        astrZeroOi.ToFile AddSlash(App.Path) & "Totals\ZeroOi.TXT"
    End If
    
    AddList "Done"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.FixTotVolAndOi"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ContractOiDiffs
'' Description: Determine the contract OI differences between for life of electronic
'' Inputs:      Electronic, Pit, Combined
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ContractOiDiffs(ByVal strElectronicBase As String, ByVal strPitBase As String, ByVal strCombinedBase As String)
On Error GoTo ErrSection:

    Dim Electronic As cGdTree           ' Collection of electronic bars
    Dim Pit As cGdTree                  ' Collection of pit bars
    Dim Combined As cGdTree             ' Collection of combined bars
    Dim ElecBars As cGdBars             ' Bars object
    Dim PitBars As cGdBars              ' Bars object
    Dim CombBars As cGdBars             ' Bars object
    Dim astrContracts As cGdArray       ' Array of futures contracts
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim astrOutput As cGdArray          ' Output array
    Dim lPitBar As Long                 ' Bar number in pit bars
    Dim lCombBar As Long                ' Bar number in combined bars
    Dim lElecOi As Long                 ' Contract OI in electronic bars
    Dim lPitOi As Long                  ' Contract OI in pit bars
    Dim lCombOi As Long                 ' Contract OI in combined bars
    Dim lSymbolID As Long               ' Symbol ID
    Dim strElecSymbol As String         ' Symbol for the electronic contract
    Dim strPitSymbol As String          ' Symbol for the pit contract
    Dim strCombSymbol As String         ' Symbol for the combined contract
    Dim astrFix As cGdArray             ' Fix file
    Dim bDumpedSymbol As Boolean        ' Have we dumped the symbol to the fix file?
    Dim astrFields As cGdArray          ' Fields to dump to the fix file

    Set Electronic = New cGdTree
    Set Pit = New cGdTree
    Set Combined = New cGdTree
    
    Set astrContracts = New cGdArray
    astrContracts.Create eGDARRAY_Strings
    
    Set astrOutput = New cGdArray
    astrOutput.Create eGDARRAY_Strings
    
    Set astrFix = New cGdArray
    astrFix.Create eGDARRAY_Strings
    
    If SU_GetContracts(GetSymbolID(strElectronicBase & "-0"), astrContracts) Then
        ' Symbol;Symbol ID ( e.g. AC-200506;77069 )
        For lIndex = 0 To astrContracts.Size - 1
            AddList "Loading '" & astrContracts(lIndex) & "'..."
            lSymbolID = CLng(Val(Parse(astrContracts(lIndex), ";", 2)))
            
            Set ElecBars = New cGdBars
            If DM_GetBars(ElecBars, lSymbolID, "Daily") Then
                Electronic.Add ElecBars, ElecBars.Prop(eBARS_Symbol)
            End If
        Next lIndex
    End If
    
    If SU_GetContracts(GetSymbolID(strPitBase & "-0"), astrContracts) Then
        ' Symbol;Symbol ID ( e.g. AC-200506;77069 )
        For lIndex = 0 To astrContracts.Size - 1
            AddList "Loading '" & astrContracts(lIndex) & "'..."
            lSymbolID = CLng(Val(Parse(astrContracts(lIndex), ";", 2)))
            
            Set PitBars = New cGdBars
            If DM_GetBars(PitBars, lSymbolID, "Daily") Then
                Pit.Add PitBars, PitBars.Prop(eBARS_Symbol)
            End If
        Next lIndex
    End If
    
    If SU_GetContracts(GetSymbolID(strCombinedBase & "-0"), astrContracts) Then
        ' Symbol;Symbol ID ( e.g. AC-200506;77069 )
        For lIndex = 0 To astrContracts.Size - 1
            AddList "Loading '" & astrContracts(lIndex) & "'..."
            lSymbolID = CLng(Val(Parse(astrContracts(lIndex), ";", 2)))
            
            Set CombBars = New cGdBars
            If DM_GetBars(CombBars, lSymbolID, "Daily") Then
                Combined.Add CombBars, CombBars.Prop(eBARS_Symbol)
            End If
        Next lIndex
    End If
    
    For lIndex = 1 To Electronic.Count
        Set ElecBars = Electronic(lIndex)
        strElecSymbol = ElecBars.Prop(eBARS_Symbol)
        strPitSymbol = Replace(strElecSymbol, strElectronicBase & "-", strPitBase & "-")
        strCombSymbol = Replace(strElecSymbol, strElectronicBase & "-", strCombinedBase & "-")
        
        If Pit.Exists(strPitSymbol) Then
            Set PitBars = Pit(strPitSymbol)
        Else
            Set PitBars = Nothing
        End If
        If Combined.Exists(strCombSymbol) Then
            Set CombBars = Combined(strCombSymbol)
        Else
            Set CombBars = Nothing
        End If
        
        bDumpedSymbol = False
        
        For lIndex2 = 0 To ElecBars.Size - 1
            AddList ElecBars.Prop(eBARS_Symbol) & " - " & DateFormat(ElecBars(eBARS_DateTime, lIndex2), MM_DD_YYYY)
            
            lElecOi = ElecBars(eBARS_ContOI, lIndex2)
            
            If PitBars Is Nothing Then
                lPitOi = kNullData
            Else
                lPitBar = PitBars.FindDateTime(ElecBars(eBARS_DateTime, lIndex2), True)
                If lPitBar = -1& Then
                    lPitOi = kNullData
                Else
                    lPitOi = PitBars(eBARS_ContOI, lPitBar)
                End If
            End If
            
            If CombBars Is Nothing Then
                lCombOi = kNullData
            Else
                lCombBar = CombBars.FindDateTime(ElecBars(eBARS_DateTime, lIndex2), True)
                If lCombBar = -1& Then
                    lCombOi = kNullData
                Else
                    lCombOi = CombBars(eBARS_ContOI, lCombBar)
                End If
            End If
            
            If (lPitOi = lCombOi) And (lPitOi <> kNullData) And (lPitOi <> lElecOi) Then
                If bDumpedSymbol = False Then
                    astrFix.Add "#" & ElecBars.Prop(eBARS_Symbol)
                    bDumpedSymbol = True
                End If
                
                Set astrFields = New cGdArray
                astrFields.Create eGDARRAY_Strings, 10
                
                astrFields(0) = Format(ElecBars(eBARS_DateTime, lIndex2), "YYYYMMDD")
                astrFields(1) = ElecBars.PriceDisplay(ElecBars(eBARS_Open, lIndex2), False)
                astrFields(2) = ElecBars.PriceDisplay(ElecBars(eBARS_High, lIndex2), False)
                astrFields(3) = ElecBars.PriceDisplay(ElecBars(eBARS_Low, lIndex2), False)
                astrFields(4) = ElecBars.PriceDisplay(ElecBars(eBARS_Close, lIndex2), False)
                astrFields(5) = Str(ElecBars(eBARS_ContVol, lIndex2))
                astrFields(6) = Str(lPitOi)
                astrFields(7) = Str(ElecBars(eBARS_Vol, lIndex2))
                astrFields(8) = Str(PitBars(eBARS_OI, lPitBar))
                astrFields(9) = "0"
                
                astrFix.Add astrFields.JoinFields(" ")
            ElseIf (lElecOi = 0) And (lPitOi <> kNullData) And (lCombOi = kNullData) Then
                If bDumpedSymbol = False Then
                    astrFix.Add "#" & ElecBars.Prop(eBARS_Symbol)
                    bDumpedSymbol = True
                End If
                
                Set astrFields = New cGdArray
                astrFields.Create eGDARRAY_Strings, 10
                
                astrFields(0) = Format(ElecBars(eBARS_DateTime, lIndex2), "YYYYMMDD")
                astrFields(1) = ElecBars.PriceDisplay(ElecBars(eBARS_Open, lIndex2), False)
                astrFields(2) = ElecBars.PriceDisplay(ElecBars(eBARS_High, lIndex2), False)
                astrFields(3) = ElecBars.PriceDisplay(ElecBars(eBARS_Low, lIndex2), False)
                astrFields(4) = ElecBars.PriceDisplay(ElecBars(eBARS_Close, lIndex2), False)
                astrFields(5) = Str(ElecBars(eBARS_ContVol, lIndex2))
                astrFields(6) = Str(lPitOi)
                astrFields(7) = Str(ElecBars(eBARS_Vol, lIndex2))
                astrFields(8) = Str(PitBars(eBARS_OI, lPitBar))
                astrFields(9) = "0"
                
                astrFix.Add astrFields.JoinFields(" ")
            ElseIf (lElecOi = 0) And (lPitOi = kNullData) And (lCombOi <> kNullData) Then
                If bDumpedSymbol = False Then
                    astrFix.Add "#" & ElecBars.Prop(eBARS_Symbol)
                    bDumpedSymbol = True
                End If
                
                Set astrFields = New cGdArray
                astrFields.Create eGDARRAY_Strings, 10
                
                astrFields(0) = Format(ElecBars(eBARS_DateTime, lIndex2), "YYYYMMDD")
                astrFields(1) = ElecBars.PriceDisplay(ElecBars(eBARS_Open, lIndex2), False)
                astrFields(2) = ElecBars.PriceDisplay(ElecBars(eBARS_High, lIndex2), False)
                astrFields(3) = ElecBars.PriceDisplay(ElecBars(eBARS_Low, lIndex2), False)
                astrFields(4) = ElecBars.PriceDisplay(ElecBars(eBARS_Close, lIndex2), False)
                astrFields(5) = Str(ElecBars(eBARS_ContVol, lIndex2))
                astrFields(6) = Str(lCombOi)
                astrFields(7) = Str(ElecBars(eBARS_Vol, lIndex2))
                astrFields(8) = Str(CombBars(eBARS_OI, lCombBar))
                astrFields(9) = "0"
                
                astrFix.Add astrFields.JoinFields(" ")
            ElseIf (lElecOi <> lPitOi) Or (lElecOi <> lCombOi) Or (lPitOi <> lCombOi) Then
                astrOutput.Add ElecBars.Prop(eBARS_Symbol) & "; " & DateFormat(ElecBars(eBARS_DateTime, lIndex2), MM_DD_YYYY) & "; Elec=" & Str(lElecOi) & "; Comb=" & Str(lCombOi) & "; Pit=" & Str(lPitOi)
            End If
        Next lIndex2
    Next lIndex
    
    astrFix.ToFile AddSlash(App.Path) & "ContOi.FIX"
    astrOutput.ToFile AddSlash(App.Path) & "ContOi.TXT"
    
    AddList "Done"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.ContractOiDiffs"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReadNifty
'' Description: Read in a Nifty file
'' Inputs:      Nifty file
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ReadNifty(ByVal strNiftyFile As String, ByVal strOutputPath As String)
On Error GoTo ErrSection:

    Dim astrNiftyFile As cGdArray       ' Nifty file read into a an array
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrFields As cGdArray          ' Fields from the line in the nifty file
    Dim Totals As cGdTree               ' Collection of total volume/open interest
    Dim strKey As String                ' Key into the collection
    Dim strTotals As String             ' Totals
    Dim strOutputFile As String         ' Output file name
    Dim astrOutput As cGdArray          ' Output file array
    Dim lPos As Long                    ' Position in the array
    Dim astrOutFields As cGdArray       ' Output fields
    Dim bAdd As Boolean                 ' Add the line to the file?
    
    Set astrNiftyFile = New cGdArray
    If astrNiftyFile.FromFile(strNiftyFile) Then
        Set astrFields = New cGdArray
        Set Totals = New cGdTree
        
        ' First pass - store up all of the total volume/open interest lines in a collection...
        AddList strNiftyFile & ": Pass1"
        For lIndex = 0 To astrNiftyFile.Size - 1
            astrFields.Clear
            astrFields.SplitFields astrNiftyFile(lIndex), " "
            
            If astrFields(0) = "FM" Then
                strKey = astrFields(1) & "|" & astrFields(2)
                strTotals = astrFields(3) & " " & astrFields(4)
                
                If Totals.Exists(strKey) Then
                    Totals(strKey) = strTotals
                Else
                    Totals.Add strTotals, strKey
                End If
            End If
        Next lIndex
        
        AddList strNiftyFile & ": Pass2"
        For lIndex = 0 To astrNiftyFile.Size - 1
            astrFields.Clear
            astrFields.SplitFields astrNiftyFile(lIndex), " "
            
            If (astrFields(0) = "FP") Or (astrFields(0) = "FV") Then
                Set astrOutput = New cGdArray
                strOutputFile = AddSlash(strOutputPath) & astrFields(1) & "\" & astrFields(2) & ".TXT"
                astrOutput.FromFile strOutputFile
                
                If DirExist(AddSlash(strOutputPath) & astrFields(1)) = False Then
                    MkDir AddSlash(strOutputPath) & astrFields(1)
                End If
                
                Set astrOutFields = New cGdArray
                
                If astrOutput.BinarySearch(astrFields(3) & " ", lPos, eGdSort_MatchUsingSearchStringLength) Then
                    bAdd = False
                    astrOutFields.SplitFields astrOutput(lPos), " "
                Else
                    bAdd = True
                    astrOutFields.Create eGDARRAY_Strings, 9
                    astrOutFields(0) = astrFields(3)
                    astrOutFields(1) = Str(kNullData)
                    astrOutFields(2) = Str(kNullData)
                    astrOutFields(3) = Str(kNullData)
                    astrOutFields(4) = Str(kNullData)
                    astrOutFields(5) = Str(kNullData)
                    astrOutFields(6) = Str(kNullData)
                    astrOutFields(7) = Str(kNullData)
                    astrOutFields(8) = Str(kNullData)
                End If
                
                If astrFields(0) = "FP" Then
                    astrOutFields(1) = astrFields(4)
                    astrOutFields(2) = astrFields(5)
                    astrOutFields(3) = astrFields(6)
                    astrOutFields(4) = astrFields(7)
                Else
                    astrOutFields(5) = astrFields(4)
                    astrOutFields(6) = astrFields(5)
                    
                    strKey = astrFields(1) & "|" & astrFields(3)
                    If Totals.Exists(strKey) Then
                        strTotals = Totals(strKey)
                        
                        astrOutFields(7) = Parse(strTotals, " ", 1)
                        astrOutFields(8) = Parse(strTotals, " ", 2)
                    End If
                End If
                
                If bAdd Then
                    astrOutput.Add astrOutFields.JoinFields(" "), lPos
                Else
                    astrOutput(lPos) = astrOutFields.JoinFields(" ")
                End If
                
                astrOutput.ToFile strOutputFile
            End If
        Next lIndex
    
        AddList strNiftyFile & ": Done"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.ReadNifty"
    
End Sub
#End If

Private Sub SecondBarFileToAsc(ByVal strSecondBarFile As String, ByVal strAscFile As String)
On Error GoTo ErrSection:

    Dim astrSecondBarFile As cGdArray   ' Second bar file split out into an array
    Dim astrAscFile As cGdArray         ' Array to write to the ASC file
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrFields As cGdArray          ' Fields in the line from the second bar file
    Dim strTime As String               ' Time for the second bar
    Dim iHour As Integer                ' Hour for the second bar
    Dim iMinute As Integer              ' Minute for the second bar
    Dim iSecond As Integer              ' Second for the second bar
    Dim iMinuteOfDay As Integer         ' Minute of the day
    
    Set astrSecondBarFile = New cGdArray
    If astrSecondBarFile.FromFile(strSecondBarFile) = True Then
        Set astrAscFile = New cGdArray
        astrAscFile.Create eGDARRAY_Strings
        
        For lIndex = 0 To astrSecondBarFile.Size - 1
            Set astrFields = New cGdArray
            astrFields.SplitFields astrSecondBarFile(lIndex), vbTab
            
            strTime = Parse(astrFields(1), " ", 2)
            iHour = Int(Val(Parse(strTime, ":", 1)))
            iMinute = Int(Val(Parse(strTime, ":", 2)))
            iSecond = Int(Val(Parse(strTime, ":", 3)))
            
            iHour = iHour + 2
            If iHour > 23 Then
                iHour = iHour - 24
            End If
            
            iMinuteOfDay = (iHour * 60) + iMinute
            
            If (iMinuteOfDay >= ((0 * 60) + 41)) And (iMinuteOfDay <= ((8 * 60) + 18)) Then
                strTime = Format(iHour, "00") & Format(iMinute, "00") & "::" & Format(iSecond, "00")
                
                astrAscFile.Add strTime & " " & astrFields(2) & " 0 1"
                If astrFields(3) <> astrFields(2) Then
                    astrAscFile.Add strTime & " " & astrFields(3) & " 0 1"
                End If
                If (astrFields(4) <> astrFields(2)) And (astrFields(4) <> astrFields(3)) Then
                    astrAscFile.Add strTime & " " & astrFields(4) & " 0 1"
                End If
                If (astrFields(5) <> astrFields(2)) And (astrFields(5) <> astrFields(3)) And (astrFields(5) <> astrFields(4)) Then
                    astrAscFile.Add strTime & " " & astrFields(5) & " 0 1"
                End If
            End If
        Next lIndex
    
        astrAscFile.ToFile strAscFile
    End If
    
    AddList "Done"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.SecondBarFileToAsc"
    
End Sub

Private Sub SplitBigAscFile(ByVal strFileName As String, ByVal strOutputPath As String)
On Error GoTo ErrSection:

    Dim astrAscFile As cGdArray         ' Big ASC file broken out into an array
    Dim astrOutputFile As cGdArray      ' Array of information to send to output file
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol
    
    Set astrAscFile = New cGdArray
    If astrAscFile.FromFile(strFileName) = True Then
        Set astrOutputFile = New cGdArray
        astrOutputFile.Create eGDARRAY_Strings
        
        For lIndex = 0 To astrAscFile.Size - 1
            If Left(astrAscFile(lIndex), 1) = "$" Then
                If lIndex > 0 Then
                    astrOutputFile.ToFile AddSlash(strOutputPath) & strSymbol & ".ASC"
                End If
                
                astrOutputFile.Clear
                astrOutputFile.Add astrAscFile(lIndex)
                strSymbol = Replace(Parse(astrAscFile(lIndex), " ", 1), "$", "")
                AddList strSymbol
            Else
                astrOutputFile.Add astrAscFile(lIndex)
            End If
        Next lIndex
    
        astrOutputFile.ToFile AddSlash(strOutputPath) & strSymbol & ".ASC"
    End If
    
    AddList "Done"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.SplitBigAscFile"
    
End Sub

Private Sub AscFileCompare(ByVal strFileName As String, ByVal strInputPath As String, ByVal strOutputFile As String)
On Error GoTo ErrSection:

    Dim astrBigAscFile As cGdArray      ' Big ASC file broken out into an array
    Dim astrAscFile As cGdArray         ' IB ASC file broken out into an array
    Dim astrOutputFile As cGdArray      ' Array of information to send to output file
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol
    
    Set astrBigAscFile = New cGdArray
    If astrBigAscFile.FromFile(strFileName) = True Then
        Set astrOutputFile = New cGdArray
        astrOutputFile.Create eGDARRAY_Strings
        
        For lIndex = 0 To astrBigAscFile.Size - 1
            If Left(astrBigAscFile(lIndex), 1) = "$" Then
                strSymbol = Replace(Parse(astrBigAscFile(lIndex), " ", 1), "$", "")
                AddList strSymbol
                
                Set astrAscFile = New cGdArray
                If astrAscFile.FromFile(AddSlash(strInputPath) & strSymbol & ".ASC") Then
                    astrOutputFile.Add astrBigAscFile(lIndex)
                    astrOutputFile.Add astrBigAscFile(lIndex + 1)
                    astrOutputFile.Add astrAscFile(0)
                End If
            End If
        Next lIndex
    
        astrOutputFile.ToFile strOutputFile
    End If
    
    AddList "Done"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.AscFileCompare"
    
End Sub

Private Sub MergeAscFiles(ByVal strSymbolFile As String, ByVal strGenesisPath As String, ByVal strIbPath As String, ByVal strOutputPath As String)
On Error GoTo ErrSection:

    Dim astrSymbols As cGdArray         ' List of symbols
    Dim astrGenesisFile As cGdArray     ' Genesis file split out into an array
    Dim astrIbFile As cGdArray          ' IB file split out into an array
    Dim astrOutputFile As cGdArray      ' Array to send to the output file
    Dim astrFields As cGdArray          ' Fields from the IB file line
    Dim strSymbol As String             ' Symbol from the file
    Dim lSymbol As Long                 ' Index into a for loop
    Dim lGenLine As Long                ' Current index into the Genesis file array
    Dim lIndex As Long                  ' Index into a for loop
    Dim iMinute As Integer              ' Minute of the day

    Set astrSymbols = New cGdArray
    If astrSymbols.FromFile(strSymbolFile) = True Then
        For lSymbol = 0 To astrSymbols.Size - 1
            strSymbol = Replace(astrSymbols(lSymbol), "$", "")
            AddList strSymbol
            
            Set astrGenesisFile = New cGdArray
            If astrGenesisFile.FromFile(AddSlash(strGenesisPath) & strSymbol & ".ASC") Then
                Set astrIbFile = New cGdArray
                If astrIbFile.FromFile(AddSlash(strIbPath) & strSymbol & ".ASC") Then
                    Set astrOutputFile = New cGdArray
                    astrOutputFile.Create eGDARRAY_Strings
                    
                    astrOutputFile.Add astrGenesisFile(0)
                    lGenLine = 1
                    Do
                        iMinute = MinuteOfTheDayForAscTime(Parse(astrGenesisFile(lGenLine), " ", 1))
                        
                        If (iMinute < ((0 * 60) + 41)) Or (iMinute >= (17 * 60) + 0) Then
                            astrOutputFile.Add astrGenesisFile(lGenLine)
                        Else
                            Exit Do
                        End If
                        
                        lGenLine = lGenLine + 1&
                    Loop While lGenLine < astrGenesisFile.Size - 1
                    
                    For lIndex = 0 To astrIbFile.Size - 1
                        Set astrFields = New cGdArray
                        astrFields.SplitFields astrIbFile(lIndex), " "
                        
                        Select Case Parse(astrGenesisFile(0), " ", 3)
                            Case "1000"
                                astrFields(1) = Format(Val(astrFields(1)), "0.000")
                            Case "10000"
                                astrFields(1) = Format(Val(astrFields(1)), "0.0000")
                            Case "100000"
                                astrFields(1) = Format(Val(astrFields(1)), "0.00000")
                        End Select
                        
                        astrOutputFile.Add astrFields.JoinFields(" ")
                    Next lIndex
                    
                    For lIndex = lGenLine To astrGenesisFile.Size - 1
                        iMinute = MinuteOfTheDayForAscTime(Parse(astrGenesisFile(lIndex), " ", 1))
                        If iMinute > ((8 * 60) + 18) Then
                            astrOutputFile.Add astrGenesisFile(lIndex)
                        End If
                    Next lIndex
                    
                    astrOutputFile.ToFile AddSlash(strOutputPath) & strSymbol & ".ASC"
                End If
            End If
        Next lSymbol
    End If
    
    AddList "Done"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.MergeAscFiles"
    
End Sub

Private Function MinuteOfTheDayForAscTime(ByVal strTime As String) As Integer
On Error GoTo ErrSection:

    Dim iReturn As Integer              ' Return value for the function
    Dim iHour As Integer                ' Hour portion of the time
    Dim iMinute As Integer              ' Minute portion of the time
    
    iReturn = -1
    If Len(strTime) = 8 Then
        If Mid(strTime, 5, 2) = "::" Then
            iHour = Int(Val(Left(strTime, 2)))
            iMinute = Int(Val(Mid(strTime, 3, 2)))
            
            iReturn = (iHour * 60) + iMinute
        End If
    End If
    
    MinuteOfTheDayForAscTime = iReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTest2.MinuteOfTheDayForAscTime"
    
End Function

Private Sub KeepOnlyMessagesReceived(ByVal strFile As String, ByVal nBroker As eTT_AccountType, Optional ByVal strNewFile = "")
On Error GoTo ErrSection:

    Dim astrFile As cGdArray            ' File broken out into an array
    Dim lIndex As Long                  ' Index into a for loop
    Dim strReceived As String           ' Message Received string
    Dim iLen As Integer                 ' Length of string
    Dim iPos As Integer                 ' Position of string in another string
    Dim iPos2 As Integer                ' Position of string in another string
    Dim strBroker As String             ' Broker name
    Dim strMessageType As String        ' Message type
    Dim strMessage As String            ' Message to send
    
    strBroker = g.Broker.BrokerName(nBroker)
    strReceived = strBroker & " Message Received"
    iLen = Len(strReceived)
    
    Set astrFile = New cGdArray
    If astrFile.FromFile(strFile) Then
        For lIndex = astrFile.Size - 1 To 0 Step -1
            iPos = InStr(astrFile(lIndex), strReceived)
            If iPos = 0 Then
                AddList "Removing: " & astrFile(lIndex)
                astrFile.Remove lIndex
            Else
                AddList "Keeping: " & astrFile(lIndex)
                
                iPos = iPos + iLen + 2
                iPos2 = InStr(iPos, astrFile(lIndex), ")")
                If iPos2 > 0 Then
                    strMessageType = Mid(astrFile(lIndex), iPos, iPos2 - iPos)
                    strMessage = Mid(astrFile(lIndex), iPos2 + 3)
                End If
                
                astrFile(lIndex) = strMessageType & "~" & strMessage
            End If
        Next lIndex
        
        If Len(strNewFile) = 0 Then
            astrFile.ToFile strFile
        Else
            astrFile.ToFile strNewFile
        End If
        AddList "Done"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.KeepOnlyMessagesReceived"
    
End Sub

Private Sub DownloaderSample()
On Error GoTo ErrSection:

    Dim DlFile As cDownloaderFile       ' Downloader file object
    
    Set m.Downloader = New cDownloader
    
    m.Downloader.Caption = "Downloader Sample"
    m.Downloader.UserName = "genuser"
    m.Downloader.Password = "Qbc951a"
    m.Downloader.IP = "199.227.139.175"
    m.Downloader.Port = 21
    m.Downloader.DoneFile = AddSlash(App.Path) & "DownloadDone.TXT"
    
    Set DlFile = New cDownloaderFile
    DlFile.ServerPath = ".\datasets\ds140711"
    DlFile.ServerFilename = "AA1_0711.gzp"
    DlFile.LocalPath = "C:\David\Daj\"
    DlFile.LocalFilename = DlFile.ServerFilename
    DlFile.IsZipFile = True
    DlFile.ZipPath = "C:\David\Daj\AA1_0711"
    m.Downloader.Files.Add DlFile

    Set DlFile = New cDownloaderFile
    DlFile.ServerPath = ".\datasets\ds140711"
    DlFile.ServerFilename = "FS7_1314.gzp"
    DlFile.LocalPath = "C:\David\Daj\"
    DlFile.LocalFilename = DlFile.ServerFilename
    DlFile.IsZipFile = True
    DlFile.ZipPath = "C:\David\Daj\FS7_1314.gzp"
    m.Downloader.Files.Add DlFile

    Set DlFile = New cDownloaderFile
    DlFile.ServerPath = ".\datasets\ds140711"
    DlFile.ServerFilename = "FS8_1213.gzp"
    DlFile.LocalPath = "C:\David\Daj\"
    DlFile.LocalFilename = DlFile.ServerFilename
    DlFile.IsZipFile = True
    DlFile.ZipPath = "C:\David\Daj\FS8_1213.gzp"
    m.Downloader.Files.Add DlFile
    
    m.Downloader.Download AddSlash(App.Path) & "DownloadInfo.TXT"
    
    tmrDownload.Interval = 5000
    tmrDownload.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTest2.DownloaderSample"
    
End Sub

