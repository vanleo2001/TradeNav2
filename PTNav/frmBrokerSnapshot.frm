VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmBrokerSnapshot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniTextBoxXP txtGenesisPrice 
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   -1  'True
      Text            =   "frmBrokerSnapshot.frx":0000
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
      Tip             =   "frmBrokerSnapshot.frx":002A
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerSnapshot.frx":004A
   End
   Begin HexUniControls.ctlUniTextBoxXP txtAskPrice 
      Height          =   315
      Left            =   1260
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   -1  'True
      Text            =   "frmBrokerSnapshot.frx":0066
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
      Tip             =   "frmBrokerSnapshot.frx":0090
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerSnapshot.frx":00B0
   End
   Begin HexUniControls.ctlUniTextBoxXP txtBidPrice 
      Height          =   315
      Left            =   1260
      TabIndex        =   4
      Top             =   960
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   -1  'True
      Text            =   "frmBrokerSnapshot.frx":00CC
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
      Tip             =   "frmBrokerSnapshot.frx":00F6
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerSnapshot.frx":0116
   End
   Begin HexUniControls.ctlUniTextBoxXP txtLastPrice 
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   600
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   -1  'True
      Text            =   "frmBrokerSnapshot.frx":0132
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
      Tip             =   "frmBrokerSnapshot.frx":015C
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerSnapshot.frx":017C
   End
   Begin HexUniControls.ctlUniComboImageXP cboSymbols 
      Height          =   315
      Left            =   900
      TabIndex        =   1
      Top             =   120
      Width           =   1695
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
      Tip             =   "frmBrokerSnapshot.frx":0198
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmBrokerSnapshot.frx":01B8
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblGenesisPrice 
      Height          =   195
      Left            =   180
      Top             =   1740
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
      Caption         =   "frmBrokerSnapshot.frx":01D4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmBrokerSnapshot.frx":0212
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerSnapshot.frx":0232
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblAskPrice 
      Height          =   195
      Left            =   180
      Top             =   1380
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
      Caption         =   "frmBrokerSnapshot.frx":024E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmBrokerSnapshot.frx":0284
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerSnapshot.frx":02A4
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblBidPrice 
      Height          =   195
      Left            =   180
      Top             =   1020
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
      Caption         =   "frmBrokerSnapshot.frx":02C0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmBrokerSnapshot.frx":02F6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerSnapshot.frx":0316
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblLastPrice 
      Height          =   195
      Left            =   180
      Top             =   660
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
      Caption         =   "frmBrokerSnapshot.frx":0332
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmBrokerSnapshot.frx":036A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerSnapshot.frx":038A
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblSymbol 
      Height          =   195
      Left            =   180
      Top             =   180
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
      Caption         =   "frmBrokerSnapshot.frx":03A6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmBrokerSnapshot.frx":03D6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmBrokerSnapshot.frx":03F6
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmBrokerSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmBrokerSnapshot.frm
'' Description: Display prices coming in from the broker
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO 80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 09/26/2011   DAJ         Clear text boxes on symbol change
'' 12/06/2011   DAJ         Added RJ O'Brien (PATS)
'' 12/13/2011   DAJ         Added Capital Trading Group for PATS
'' 07/22/2014   DAJ         Added Demo PATS
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    nBroker As eTT_AccountType          ' Broker for this form
    strSymbol As String                 ' Symbol currently subscribed to
End Type
Private m As mPrivate

Public Property Get Broker() As eTT_AccountType
    Broker = m.nBroker
End Property

Public Property Get Symbol() As String
    Symbol = cboSymbols.Text
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:
    
    m.nBroker = nBroker
    Caption = g.Broker.BrokerName(nBroker) & " Data"
    
    LoadSymbolsCombo
    cboSymbols.ListIndex = -1&

    ShowForm Me, eForm_Nonmodal, frmMain

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerSnapshot.ShowMe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Broker_DataUpdate
'' Description: Handle a data update from the broker
'' Inputs:      Symbol, Last, Bid, Ask
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Broker_DataUpdate(ByVal strSymbol As String, Optional ByVal dLast As Double = kNullData, Optional ByVal dBid As Double = kNullData, Optional ByVal dAsk As Double = kNullData)
On Error GoTo ErrSection:

    If strSymbol = m.strSymbol Then
        If dLast <> kNullData Then
            txtLastPrice.Text = Str(dLast)
        End If
        If dBid <> kNullData Then
            txtBidPrice.Text = Str(dBid)
        End If
        If dAsk <> kNullData Then
            txtAskPrice.Text = Str(dAsk)
        End If
        
        txtGenesisPrice.Text = g.RealTime.LastKnownPrice(m.strSymbol)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerSnapshot.Broker_DataUpdate", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboSymbols_Click
'' Description: Handle the user changing symbols
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboSymbols_Click()
On Error GoTo ErrSection:

    If Len(m.strSymbol) > 0 Then
        Unsubscribe
    End If
    
    txtAskPrice.Text = ""
    txtBidPrice.Text = ""
    txtLastPrice.Text = ""
    txtGenesisPrice.Text = ""
    
    Subscribe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerSnapshot.cboSymbols_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    CenterTheForm Me
    Icon = Picture16("kBlank")
    
    g.Styler.StyleForm Me

    txtLastPrice.Text = ""
    txtBidPrice.Text = ""
    txtAskPrice.Text = ""
    txtGenesisPrice.Text = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerSnapshot.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Unsubscribe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerSnapshot.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Symbols
'' Description: Symbols object from the broker
'' Inputs:      None
'' Returns:     Symbols object (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Symbols() As cBrokerSymbols
On Error GoTo ErrSection:

    Dim returnSymbols As cBrokerSymbols ' Broker symbols object to return
    
    Set returnSymbols = Nothing
    Select Case m.nBroker
        Case eTT_AccountType_CtgPats
            Set returnSymbols = g.CtgPats.Symbols
            
        Case eTT_AccountType_DemoPats
            Set returnSymbols = g.DemoPats.Symbols
            
        Case eTT_AccountType_RjoPats
            Set returnSymbols = g.RjoPats.Symbols
            
        Case eTT_AccountType_TT
            Set returnSymbols = g.TT.Symbols
            
    End Select
    
    Set Symbols = returnSymbols

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBrokerSnapshot.Symbols"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadSymbolsCombo
'' Description: Load the symbols combo
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadSymbolsCombo()
On Error GoTo ErrSection:

    Dim BrokerSymbols As cBrokerSymbols ' Broker symbols
    Dim BrokerSymbol As cBrokerSymbol   ' Broker symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbol As String             ' Symbol to add to the combo
    Dim dCurrentTime As Double          ' Current time
    
    Set BrokerSymbols = Symbols
    With cboSymbols
        .Clear
        
        dCurrentTime = CurrentTime
        For lIndex = 1 To BrokerSymbols.Count
            Set BrokerSymbol = BrokerSymbols(lIndex)
            If (Left(BrokerSymbol.BrokerBase, 2) <> "O:") And (BrokerSymbol.BrokerBase <> "!") And (BrokerSymbol.BrokerBase <> "@") Then
                strSymbol = ConvertToTradeSymbol(BrokerSymbol.GenesisBase & "-067", dCurrentTime)
                .AddItem strSymbol
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerSnapshot.LoadSymbolsCombo"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Subscribe
'' Description: Subscribe to the symbol in the combo box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Subscribe()
On Error GoTo ErrSection:

    Dim strSymbol As String             ' Symbol to subscribe to

    strSymbol = cboSymbols.Text
    If Len(strSymbol) > 0 Then
        Select Case m.nBroker
            Case eTT_AccountType_CtgPats
                If Not g.CtgPats Is Nothing Then
                    g.CtgPats.Subscribe strSymbol
                End If
                
            Case eTT_AccountType_DemoPats
                If Not g.DemoPats Is Nothing Then
                    g.DemoPats.Subscribe strSymbol
                End If
                
            Case eTT_AccountType_RjoPats
                If Not g.RjoPats Is Nothing Then
                    g.RjoPats.Subscribe strSymbol
                End If
                
            Case eTT_AccountType_TT
                If Not g.TT Is Nothing Then
                    g.TT.Subscribe strSymbol
                End If
                
        End Select
        
        m.strSymbol = strSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerSnapshot.Subscribe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Unsubscribe
'' Description: Unsubscribe from the previous symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Unsubscribe()
On Error GoTo ErrSection:

    If Len(m.strSymbol) > 0 Then
        Select Case m.nBroker
            Case eTT_AccountType_CtgPats
                If Not g.CtgPats Is Nothing Then
                    g.CtgPats.Unsubscribe m.strSymbol
                End If
                
            Case eTT_AccountType_DemoPats
                If Not g.DemoPats Is Nothing Then
                    g.DemoPats.Unsubscribe m.strSymbol
                End If
                
            Case eTT_AccountType_RjoPats
                If Not g.RjoPats Is Nothing Then
                    g.RjoPats.Unsubscribe m.strSymbol
                End If
                
            Case eTT_AccountType_TT
                If Not g.TT Is Nothing Then
                    g.TT.Unsubscribe m.strSymbol
                End If
        
        End Select
        
        m.strSymbol = ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBrokerSnapshot.Unsubscribe"
    
End Sub

