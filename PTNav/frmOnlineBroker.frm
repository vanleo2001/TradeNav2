VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmOnlineBroker 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdateCharts 
      Left            =   6060
      Top             =   180
   End
   Begin VB.Timer tmrAutoTradeAction 
      Enabled         =   0   'False
      Left            =   5580
      Top             =   180
   End
   Begin VB.Timer tmrAutoTradeData 
      Enabled         =   0   'False
      Left            =   5100
      Top             =   180
   End
   Begin VB.Timer tmrGmaj 
      Left            =   4620
      Top             =   180
   End
   Begin HexUniControls.ctlUniRichTextBoxXP rtbRichText 
      Height          =   435
      Left            =   180
      TabIndex        =   2
      Top             =   1200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   767
      BackColor       =   -2147483643
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmOnlineBroker.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   -1
      MultiLine       =   -1  'True
      Alignment       =   0
      ScrollBars      =   3
      PasswordChar    =   ""
      TrapTab         =   0   'False
      RaiseChangeEvent=   -1  'True
      RaiseUpdateEvent=   0   'False
      RaiseSelChangeEvent=   -1  'True
      Tip             =   "frmOnlineBroker.frx":0020
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOnlineBroker.frx":0040
      ViewMode        =   0
      TextModeText    =   2
      TextModeUndoLevel=   8
      TextModeCodePage=   32
      AutoURLDetect   =   0   'False
      FileName        =   ""
      VerticalLayout  =   0   'False
      OnlyNumbers     =   0   'False
      NoIME           =   0   'False
      SelfIME         =   0   'False
      LanguageOptions =   150
      RaiseRequestResizeEvent=   0   'False
      RaiseMsgFilterEvent=   0   'False
      SubClassPaintMessage=   0   'False
      TabSize         =   4
      TypographyOptions=   0
      BlockAutoCopy   =   0   'False
      BlockAutoCut    =   0   'False
      BlockAutoPaste  =   0   'False
      BlockAutoUndo   =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtSalmonCallbackTs 
      Height          =   315
      Left            =   4860
      TabIndex        =   1
      Top             =   780
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmOnlineBroker.frx":005C
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
      Tip             =   "frmOnlineBroker.frx":009C
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOnlineBroker.frx":00BC
   End
   Begin VB.Timer tmrTradeServer 
      Enabled         =   0   'False
      Left            =   4140
      Top             =   180
   End
   Begin VB.Timer tmrDlgMessages 
      Left            =   3660
      Top             =   180
   End
   Begin VB.Timer tmrDanielCode 
      Left            =   3180
      Top             =   180
   End
   Begin HexUniControls.ctlUniTextBoxXP txtSalmonCallback 
      Height          =   315
      Left            =   3480
      TabIndex        =   0
      Top             =   780
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmOnlineBroker.frx":00D8
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
      Tip             =   "frmOnlineBroker.frx":0114
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOnlineBroker.frx":0134
   End
   Begin VB.Timer tmrHeartbeat 
      Enabled         =   0   'False
      Left            =   2700
      Top             =   180
   End
   Begin VB.Timer tmrMessages 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2220
      Top             =   180
   End
   Begin gdOCX.gdAppMail gdBroker 
      Left            =   180
      Top             =   180
      _ExtentX        =   953
      _ExtentY        =   847
      ControlName     =   "TNBroker"
   End
   Begin gdOCX.gdAppMail apmOptNav 
      Left            =   840
      Top             =   180
      _ExtentX        =   953
      _ExtentY        =   847
      ControlName     =   "OptNavTN"
   End
   Begin VB.Image imgRed 
      Height          =   195
      Left            =   240
      Picture         =   "frmOnlineBroker.frx":0150
      Top             =   840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgYellow 
      Height          =   195
      Left            =   540
      Picture         =   "frmOnlineBroker.frx":03D6
      Top             =   840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgGreen 
      Height          =   195
      Left            =   840
      Picture         =   "frmOnlineBroker.frx":065C
      Top             =   840
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "frmOnlineBroker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmOnlineBroker.frm
'' Description: Form to hold the OCX objects for the Online Broker's
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Date         Author      Description
'' 04/20/2009   DAJ         Added support for option chain structure calls
'' 05/19/2009   DAJ         Use the new Option Navigator status variable
'' 06/05/2009   DAJ         Added ConnectToAccount call from Option Navigator
'' 06/22/2009   DAJ         Added TicketSubmitted message
'' 09/28/2009   TLB         Moved MessageReceived code to mOptNav
'' 10/07/2009   DAJ         Added support for Linked Orders at broker
'' 03/11/2010   DAJ         Added the red, green, and yellow icons
'' 04/20/2010   DAJ         Set the AutoSendInterval on OptNav app mail object
'' 07/20/2010   DAJ         Added the Daniel Code timer
'' 07/21/2010   DAJ         Changed over to the Provided.INI for Daniel Code stuff
'' 08/05/2010   DAJ         Disable DanielCode button when stand alone is running
'' 08/09/2010   DAJ         Added dialog message queue
'' 09/13/2010   DAJ         Added code for Rithmic
'' 10/26/2010   DAJ         Dump entire message queue, smart queue for Rithmic
'' 11/12/2010   DAJ         Added Optimus, OpVest, and Vision
'' 12/10/2010   DAJ         Added Zen-Fire, Changed over to the IsBrokerUser function
'' 03/07/2011   DAJ         Added OEC/Options Express, IB/I-Deal/Rithmic/Gain to broker base class
'' 05/11/2011   DAJ         Utilize IsLiveAccount function
'' 06/21/2011   DAJ         Separate out Simulated trading types
'' 08/25/2011   DAJ         Moved some code to cBroker, mods for CQG/TT
'' 11/02/2011   DAJ         Added Amp Trading and RJ O'Brien as CQG brokers
'' 12/02/2101   DAJ         Added RJO (PATS)
'' 12/09/2011   DAJ         Added GFT Forex and OptionsHouse
'' 12/13/2011   DAJ         Added Capital Trading Group for PATS and CQG
'' 12/14/2011   DAJ         Added Capital Trading Group and Fintec for PFG
'' 03/14/2012   DAJ         Added Alpari(Currenex), Alpari(PATS), Penson(Currenex), Penson(CQG)
'' 05/31/2012   DAJ         Turnkey implementation
'' 06/07/2012   DAJ         Clean up Turnkey upon unload
'' 07/16/2012   DAJ         ZanerCqg, ZanerPats, ZanerRithmic, ZanerZenFire, KnightCnx, KnightCqg
'' 07/17/2012   DAJ         AlpariZenFire
'' 07/17/2012   DAJ         RobbinsCqg
'' 07/18/2012   DAJ         RCG (New PATS)
'' 07/27/2012   DAJ         Demo (PATS), GmajPro
'' 08/02/2012   DAJ         If GmajPro or DanielCode are done, re-enable both toolbar buttons
'' 08/03/2012   DAJ         Remove Gain, FXCM, Photon, OptionsHouse, Alaron, Cadent, Lotus, OptXpress, Oec, Robbins
'' 08/21/2012   DAJ         Rename 'GmajPro' to 'DC Genie Pro' and add icons
'' 08/23/2012   DAJ         Born (PATS), RJO Hong Kong (PATS)
'' 08/29/2012   DAJ         Zaner (Currenex)
'' 08/31/2012   DAJ         Load different INI file properties for GmajPro
'' 09/12/2012   DAJ         Added Currenex, FXDD (Currenex), and VanKar (Currenex)
'' 09/12/2012   DAJ         Removed Rosenthal (Old PATS), Changed Generic PATS to New PATS
'' 10/30/2012   DAJ         Added logging for the broker message queue
'' 12/11/2012   DAJ         Vision (CQG), Send TransAct Connection Info direct, don't show dialog
''                          if processing app-mail or message queue
'' 01/07/2013   DAJ         Profiling for trade stuff ( for Brady and Tim )
'' 05/10/2013   DAJ         Broker override list for E_SWIZX
'' 05/29/2013   DAJ         Automated Trading Timers
'' 06/24/2013   DAJ         Timer Logging
'' 07/30/2013   DAJ         Automatic journal for a fill
'' 10/16/2013   DAJ         Removed PFG/Xpress/OrderLinks, Added Oec/FptOec/FptCqg
'' 11/15/2013   DAJ         Moved turnkey object initialization to mMain.Main
'' 02/25/2014   DAJ         Fix for "Object variable..." error in CheckFillsToJournal
'' 03/07/2014   DAJ         Moved Cattle stuff into NavCattle.DLL
'' 04/22/2014   DAJ         No longer login to TransAct as "simuser"
'' 08/08/2014   DAJ         Reworked enabled symbols list for TransAct
'' 08/22/2014   DAJ         Added E-Trade
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 09/14/2015   DAJ         Added Tradier
'' 03/18/2016   DAJ         Added TD Ameritrade
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Messages As cGdTree                 ' Collection of queued up messages
    astrDlgMessages As cGdArray         ' Dialog messages to show
    
    bProcessingAppMail As Boolean       ' Are we currently processing an App Mail message?
    bProcessingMessageQueue As Boolean  ' Are we currenlty processing the message queue?
    
    FillsToJournal As cGdTree           ' Collection of fills to journal after charts are updated
End Type
Private m As mPrivate

Public Property Get ProcessingAppMail() As Boolean
    ProcessingAppMail = m.bProcessingAppMail
End Property

Public Property Get ProcessingMessageQueue() As Boolean
    ProcessingMessageQueue = m.bProcessingMessageQueue
End Property

Public Property Get FillsToJournal() As cGdTree
    Set FillsToJournal = m.FillsToJournal
End Property

Private Property Get MessageTimerEnabled() As Boolean
    MessageTimerEnabled = tmrMessages.Enabled
End Property
Private Property Let MessageTimerEnabled(ByVal bEnabled As Boolean)
    If tmrMessages.Enabled <> bEnabled Then
        DumpDebug "MessageTimerEnabled changed from " & Str(tmrMessages.Enabled) & " to " & Str(bEnabled)
    End If
    tmrMessages.Enabled = bEnabled
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddDialogMessage
'' Description: Add a dialog message to be shown with a timer
'' Inputs:      Message, Caption, Icon, Buttons
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddDialogMessage(ByVal strMessage As String, Optional ByVal strIcon As String = "", Optional ByVal strButtons As String = "", Optional ByVal strCaption As String = "")
On Error GoTo ErrSection:

    m.astrDlgMessages.Add strMessage & vbTab & strIcon & vbTab & strButtons & vbTab & strCaption
    If tmrDlgMessages.Enabled = False Then
        tmrDlgMessages.Enabled = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.AddDialogMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    apmOptNav_MessageReceived
'' Description: Handle a message received from Option Navigator
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub apmOptNav_MessageReceived(msg As gdOCX.gdAppMailMsg)
On Error GoTo ErrSection:

    apmOptNav.Tag = msg.FromControlName
    
    OptNavMessageReceived msg
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmMain.apmOptNav_MessageReceived"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    g.Styler.StyleForm Me
    
    If Not DirExist(AddSlash(App.Path) & "Brokers") Then
        MkDir AddSlash(App.Path) & "Brokers"
    End If

    Set m.Messages = New cGdTree
    Set m.astrDlgMessages = New cGdArray
    m.astrDlgMessages.Create eGDARRAY_Strings
    
    ' When we start up, assume that Option Navigator is not loaded until it tells
    ' us it is loaded...
    g.nOptNavStatus = eGDOptNavStatus_Unloaded

    ' Make sure that the OptNav path exists for dumping logs to and clean up
    ' any logs older than 30 days from the path...
    If Not DirExist(AddSlash(App.Path) & "OptNav") Then MkDir AddSlash(App.Path) & "OptNav"
    KillFile AddSlash(App.Path) & "OptNav\*.LOG /o=-30"
    
    g.Broker.InitBrokerObjects
    GetLastKnownBrokerSymbols

    Me.Visible = False
    
    ' Load up the order link information...
'    Set g.OrderLinks = New cOrderLinks
'    g.OrderLinks.Load
    
    ' Activate the Broker and Option Navigator AppMail communication...
    gdBroker.Active = True
    
    apmOptNav.AutoSendInterval = 1000
    apmOptNav.Active = True
    
    tmrHeartbeat.Interval = 5 * 1000
    tmrHeartbeat.Enabled = True
    
    tmrMessages.Interval = 5
    MessageTimerEnabled = False
    
    tmrDanielCode.Interval = 1000
    tmrDanielCode.Enabled = False
        
    tmrGmaj.Interval = 1000
    tmrGmaj.Enabled = False
        
    tmrDlgMessages.Interval = 100
    tmrDlgMessages.Enabled = False
    
    tmrTradeServer.Interval = 100
    tmrTradeServer.Enabled = False
    
    tmrAutoTradeAction.Interval = 100
    tmrAutoTradeAction.Enabled = False
    
    tmrUpdateCharts.Interval = 100
    tmrUpdateCharts.Enabled = False
    
    m.bProcessingAppMail = False
    m.bProcessingMessageQueue = False
    
    Set m.FillsToJournal = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up after ourseleves
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    ' Tell Option Navigator to unload...
    SendMessageToOptNav eGDOptNav_Unload, "TradeNav is unloading"
    
    g.Broker.DestroyBrokerObjects
    
    g.CattleBridge.Shutdown
    
    gdBroker.Unload
    apmOptNav.Unload
    
    tmrHeartbeat.Enabled = False
    MessageTimerEnabled = False
    tmrDanielCode.Enabled = False
    tmrGmaj.Enabled = False
    tmrDlgMessages.Enabled = False
    tmrAutoTradeAction.Enabled = False
    tmrAutoTradeData.Enabled = False
    tmrUpdateCharts.Enabled = False
    
    Set m.Messages = Nothing
    Set m.astrDlgMessages = Nothing
    Set m.FillsToJournal = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    gdBroker_MessageReceived
'' Description: Handle an incoming message from the TransAct program
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub gdBroker_MessageReceived(msg As gdOCX.gdAppMailMsg)
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message received
    Dim astrMessage As New cGdArray     ' Message broken out into an array
    Dim strKey As String                ' Key into the collection

    m.bProcessingAppMail = True
    strMessage = msg.Message

    Select Case UCase(msg.FromControlName)
        Case "ADVANTAGEFUTURES"
            If Not g.AdvFut Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.AdvFut.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
        
        Case "ALPARICNX"
            If Not g.AlpariCnx Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.AlpariCnx.Broker.HandleMessage msg.MsgType, strMessage
                            
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "ALPARIPATS"
            If Not g.AlpariPats Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.AlpariPats.Broker.HandleMessage msg.MsgType, strMessage
                            
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "ALPARIZENFIRE"
            If Not g.AlpariZenFire Is Nothing Then
                Select Case msg.MsgType
                    Case eGDRithmicMessageType_ConnectionInfo, eGDRithmicMessageType_AppLoaded, _
                            eGDRithmicMessageType_AppUnloaded, eGDRithmicMessageType_Heartbeat
                        g.AlpariZenFire.HandleMessage msg.MsgType, strMessage
                        
                    Case eGDRithmicMessageType_Order
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(5)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                        
                    Case eGDRithmicMessageType_Position
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(1)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                    
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
            
        Case "AMERITRADE"
            If Not g.Ameritrade Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, eGDBrokerMessageType_AppUnloaded
                        g.Ameritrade.Broker.HandleMessage msg.MsgType, strMessage
                            
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
            
        Case "AMPCQG"
            If Not g.AmpCqg Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.AmpCqg.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "BORNPATS"
            If Not g.BornPats Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.BornPats.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "CQG"
            If Not g.CQG Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.CQG.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "CTGCQG"
            If Not g.CtgCqg Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.CtgCqg.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "CTGPATS"
            If Not g.CtgPats Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.CtgPats.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
'        Case "CTGPFG"
'            If Not g.CtgPfg Is Nothing Then
'                Select Case msg.MsgType
'                    Case eGDPfgMessageType_ConnectionInfo, eGDPfgMessageType_AppLoaded, _
'                            eGDPfgMessageType_AppUnloaded, eGDPfgMessageType_Heartbeat
'                        g.CtgPfg.HandleMessage msg.MsgType, strMessage
'
'                    Case Else
'                        AddMessageToQueue msg
'
'                End Select
'            End If
            
        Case "CURRENEX"
            If Not g.Currenex Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.Currenex.Broker.HandleMessage msg.MsgType, strMessage
                            
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "DEMOPATS"
            If Not g.DemoPats Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.DemoPats.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "ETRADE"
            If Not g.Etrade Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded, eGDBrokerMessageType_LoginUrl
                        g.Etrade.Broker.HandleMessage msg.MsgType, strMessage
                            
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
            
'        Case "FINTECPFG"
'            If Not g.FintecPfg Is Nothing Then
'                Select Case msg.MsgType
'                    Case eGDPfgMessageType_ConnectionInfo, eGDPfgMessageType_AppLoaded, _
'                            eGDPfgMessageType_AppUnloaded, eGDPfgMessageType_Heartbeat
'                        g.FintecPfg.HandleMessage msg.MsgType, strMessage
'
'                    Case Else
'                        AddMessageToQueue msg
'
'                End Select
'            End If
            
        Case "FPTCQG"
            If Not g.FptCqg Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.FptCqg.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "FPTOEC"
            If Not g.FptOec Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.FptOec.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "FXDDCNX"
            If Not g.FxddCnx Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.FxddCnx.Broker.HandleMessage msg.MsgType, strMessage
                            
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "GFT"
            If Not g.Gft Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.Gft.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "IDEAL"
            If Not g.Ideal Is Nothing Then
                Select Case msg.MsgType
                    Case eGDIbMessageType_ConnectionInfo, eGDIbMessageType_AppLoaded, _
                            eGDIbMessageType_AppUnloaded, eGDIbMessageType_Heartbeat
                        g.Ideal.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
            
        Case "INTERACTIVEBROKERS"
            If Not g.IntBroker Is Nothing Then
                Select Case msg.MsgType
                    Case eGDIbMessageType_ConnectionInfo, eGDIbMessageType_AppLoaded, _
                            eGDIbMessageType_AppUnloaded, eGDIbMessageType_Heartbeat
                        g.IntBroker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
            
        Case "KNIGHTCNX"
            If Not g.KnightCnx Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.KnightCnx.Broker.HandleMessage msg.MsgType, strMessage
                            
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "KNIGHTCQG"
            If Not g.KnightCqg Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.KnightCqg.Broker.HandleMessage msg.MsgType, strMessage
                            
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
'        Case "LINDWALDOCK"
'            If Not g.LindWaldock Is Nothing Then
'                Select Case msg.MsgType
'                    Case eGDLindXpressMessageType_ConnectionInfo, eGDLindXpressMessageType_AppLoaded, _
'                            eGDLindXpressMessageType_AppUnloaded
'                        g.LindWaldock.HandleMessage msg.MsgType, strMessage
'
'                    Case Else
'                        AddMessageToQueue msg
'
'                End Select
'            End If
'
'        Case "MANEXPRESS"
'            If Not g.ManExpress Is Nothing Then
'                Select Case msg.MsgType
'                    Case eGDLindXpressMessageType_ConnectionInfo, eGDLindXpressMessageType_AppLoaded, _
'                            eGDLindXpressMessageType_AppUnloaded
'                        g.ManExpress.HandleMessage msg.MsgType, strMessage
'
'                    Case Else
'                        AddMessageToQueue msg
'
'                End Select
'            End If
            
        Case "OEC"
            If Not g.Oec Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.Oec.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "OPTIMUS"
            If Not g.Optimus Is Nothing Then
                Select Case msg.MsgType
                    Case eGDRithmicMessageType_ConnectionInfo, eGDRithmicMessageType_AppLoaded, _
                            eGDRithmicMessageType_AppUnloaded, eGDRithmicMessageType_Heartbeat
                        g.Optimus.HandleMessage msg.MsgType, strMessage
                        
                    Case eGDRithmicMessageType_Order
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(5)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                        
                    Case eGDRithmicMessageType_Position
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(1)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                    
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
            
        Case "OPVEST"
            If Not g.OpVest Is Nothing Then
                Select Case msg.MsgType
                    Case eGDRithmicMessageType_ConnectionInfo, eGDRithmicMessageType_AppLoaded, _
                            eGDRithmicMessageType_AppUnloaded, eGDRithmicMessageType_Heartbeat
                        g.OpVest.HandleMessage msg.MsgType, strMessage
                        
                    Case eGDRithmicMessageType_Order
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(5)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                        
                    Case eGDRithmicMessageType_Position
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(1)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                    
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
            
        Case "PATS"
            If Not g.Pats Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.Pats.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
'        Case "PFG"
'            If Not g.PFG Is Nothing Then
'                Select Case msg.MsgType
'                    Case eGDPfgMessageType_ConnectionInfo, eGDPfgMessageType_AppLoaded, _
'                            eGDPfgMessageType_AppUnloaded, eGDPfgMessageType_Heartbeat
'                        g.PFG.HandleMessage msg.MsgType, strMessage
'
'                    Case Else
'                        AddMessageToQueue msg
'
'                End Select
'            End If
            
        Case "RCGPATS"
            If Not g.RcgPats Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.RcgPats.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "RITHMIC"
            If Not g.Rithmic Is Nothing Then
                Select Case msg.MsgType
                    Case eGDRithmicMessageType_ConnectionInfo, eGDRithmicMessageType_AppLoaded, _
                            eGDRithmicMessageType_AppUnloaded, eGDRithmicMessageType_Heartbeat
                        g.Rithmic.HandleMessage msg.MsgType, strMessage
                        
                    Case eGDRithmicMessageType_Order
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(5)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                        
                    Case eGDRithmicMessageType_Position
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(1)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                    
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
            
        Case "RJOCQG"
            If Not g.RjoCqg Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.RjoCqg.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "RJOPATS"
            If Not g.RjoPats Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.RjoPats.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "RJOHKPATS"
            If Not g.RjoHkPats Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.RjoHkPats.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "ROBBINSCQG"
            If Not g.RobbinsCqg Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.RobbinsCqg.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "TRADIER"
            If Not g.Tradier Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded, eGDBrokerMessageType_LoginUrl
                        g.Tradier.Broker.HandleMessage msg.MsgType, strMessage
                            
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
            
        Case "TRANSACT"
            If Not g.Transact Is Nothing Then
                Select Case msg.MsgType
                    Case eGDTransActMessageType_Disconnected, eGDTransActMessageType_Subscribed, _
                            eGDTransActMessageType_Unsubscribed, eGDTransActMessageType_AppLoaded, _
                            eGDTransActMessageType_AppUnloaded, eGDTransActMessageType_Heartbeat, _
                            eGDTransActMessageType_PriceUpdate, eGDTransActMessageType_ConnectionInfo
                        g.Transact.HandleMessage msg.MsgType, strMessage
                        
                    Case Else
                        AddMessageToQueue msg
                
                End Select
            End If
            
        Case "TRADINGTECHNOLOGIES"
            If Not g.TT Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.TT.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
            
        Case "VANKARCNX"
            If Not g.VanKarCnx Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.VanKarCnx.Broker.HandleMessage msg.MsgType, strMessage
                            
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "VISION"
            If Not g.Vision Is Nothing Then
                Select Case msg.MsgType
                    Case eGDRithmicMessageType_ConnectionInfo, eGDRithmicMessageType_AppLoaded, _
                            eGDRithmicMessageType_AppUnloaded, eGDRithmicMessageType_Heartbeat
                        g.Vision.HandleMessage msg.MsgType, strMessage
                        
                    Case eGDRithmicMessageType_Order
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(5)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                        
                    Case eGDRithmicMessageType_Position
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(1)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                    
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
            
        Case "VISIONCQG"
            If Not g.VisionCqg Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.VisionCqg.Broker.HandleMessage msg.MsgType, strMessage
                            
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "ZANERCNX"
            If Not g.ZanerCnx Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.ZanerCnx.Broker.HandleMessage msg.MsgType, strMessage
                            
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "ZANERCQG"
            If Not g.ZanerCqg Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.ZanerCqg.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "ZANERPATS"
            If Not g.ZanerPats Is Nothing Then
                Select Case msg.MsgType
                    Case eGDBrokerMessageType_ConnectionInfo, eGDBrokerMessageType_AppLoaded, _
                            eGDBrokerMessageType_AppUnloaded
                        g.ZanerPats.Broker.HandleMessage msg.MsgType, strMessage
                    
                    Case Else
                        AddMessageToQueue msg

                End Select
            End If
            
        Case "ZANERRITHMIC"
            If Not g.ZanerRithmic Is Nothing Then
                Select Case msg.MsgType
                    Case eGDRithmicMessageType_ConnectionInfo, eGDRithmicMessageType_AppLoaded, _
                            eGDRithmicMessageType_AppUnloaded, eGDRithmicMessageType_Heartbeat
                        g.ZanerRithmic.HandleMessage msg.MsgType, strMessage
                        
                    Case eGDRithmicMessageType_Order
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(5)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                        
                    Case eGDRithmicMessageType_Position
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(1)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                    
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
            
        Case "ZANERZENFIRE"
            If Not g.ZanerZenFire Is Nothing Then
                Select Case msg.MsgType
                    Case eGDRithmicMessageType_ConnectionInfo, eGDRithmicMessageType_AppLoaded, _
                            eGDRithmicMessageType_AppUnloaded, eGDRithmicMessageType_Heartbeat
                        g.ZanerZenFire.HandleMessage msg.MsgType, strMessage
                        
                    Case eGDRithmicMessageType_Order
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(5)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                        
                    Case eGDRithmicMessageType_Position
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(1)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                    
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
            
        Case "ZENFIRE"
            If Not g.ZenFire Is Nothing Then
                Select Case msg.MsgType
                    Case eGDRithmicMessageType_ConnectionInfo, eGDRithmicMessageType_AppLoaded, _
                            eGDRithmicMessageType_AppUnloaded, eGDRithmicMessageType_Heartbeat
                        g.ZenFire.HandleMessage msg.MsgType, strMessage
                        
                    Case eGDRithmicMessageType_Order
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(5)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                        
                    Case eGDRithmicMessageType_Position
                        astrMessage.SplitFields strMessage, vbTab
                        If Len(astrMessage(1)) > 0 Then
                            strKey = msg.FromControlName & "|" & Str(msg.MsgType) & "|" & astrMessage(0) & "|" & astrMessage(1) & "|" & astrMessage(2)
                            AddMessageToQueue msg, strKey
                        End If
                    
                    Case Else
                        AddMessageToQueue msg
                        
                End Select
            End If
            
    End Select

ErrExit:
    m.bProcessingAppMail = False
    Exit Sub
    
ErrSection:
    m.bProcessingAppMail = False
    RaiseError "frmOnlineBroker.gdBroker_MessageReceived"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrAutoTradeAction_Timer
'' Description: Timer to perform actions for automated trading items
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrAutoTradeAction_Timer()
On Error GoTo ErrSection:

    TimerStart "frmOnlineBroker.tmrAutoTradeAction"
    If Not g.TradingItems Is Nothing Then
        g.TradingItems.DoActionChecks
    End If
    TimerEnd "frmOnlineBroker.tmrAutoTradeAction", tmrAutoTradeAction.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.tmrAutoTradeAction_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrAutoTradeData_Timer
'' Description: Timer to update the streaming data for automated trading items
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrAutoTradeData_Timer()
On Error GoTo ErrSection:

    TimerStart "frmOnlineBroker.tmrAutoTradeData"
    If Not g.TradingItems Is Nothing Then
        g.TradingItems.UpdateBars
    End If
    TimerEnd "frmOnlineBroker.tmrAutoTradeData", tmrAutoTradeData.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.tmrAutoTradeData_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrDanielCode_Timer
'' Description: Check for the file that the Daniel Code program outputs
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrDanielCode_Timer()
On Error GoTo ErrSection:

    Dim astrOrders As cGdArray          ' Orders from the file
    Dim strDcFile As String             ' Daniel code file

    TimerStart "frmOnlineBroker.tmrDanielCode"
    strDcFile = GetIniFileProperty("OutputFile", "", "DanielCode", AddSlash(App.Path) & "Provided\Provided.INI")
    If Len(strDcFile) = 0 Then
        tmrDanielCode.Enabled = False
    Else
        If FileExist(AddSlash(App.Path) & strDcFile) Then
            tmrDanielCode.Enabled = False
            
            Set astrOrders = New cGdArray
            astrOrders.FromFile AddSlash(App.Path) & strDcFile
            KillFile AddSlash(App.Path) & strDcFile
            
            If astrOrders.Size > 0 Then
                frmDanielConfirmation.ShowMe astrOrders, False
            End If
        
            frmMain.tbToolbar.Tools("ID_DanCodeWeb").Enabled = True
            frmMain.tbToolbar.Tools("ID_GmajPro").Enabled = True
        End If
    End If
    TimerEnd "frmOnlineBroker.tmrDanielCode", tmrDanielCode.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.tmrDanielCode_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrDlgMessages_Timer
'' Description: Show a dialog with a message from the message array
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrDlgMessages_Timer()
On Error GoTo ErrSection:

    Dim astrMessage As cGdArray         ' Message broken out into pieces

    TimerStart "frmOnlineBroker.tmrDlgMessages"
    If (m.bProcessingAppMail = False) And (m.bProcessingMessageQueue = False) Then
        If m.astrDlgMessages.Size > 0 Then
            tmrDlgMessages.Enabled = False
            
            Set astrMessage = New cGdArray
            astrMessage.SplitFields m.astrDlgMessages(0), vbTab
            InfBox astrMessage(0), astrMessage(1), astrMessage(2), astrMessage(3)
            m.astrDlgMessages.Remove 0
            
            If m.astrDlgMessages.Size > 0 Then
                tmrDlgMessages.Enabled = True
            End If
        End If
    End If
    TimerEnd "frmOnlineBroker.tmrDlgMessages", tmrDlgMessages.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.tmrDlgMessages_Timer"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrGmaj_Timer
'' Description: Check for the file that the GmajPro program outputs
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrGmaj_Timer()
On Error GoTo ErrSection:

    Dim astrOrders As cGdArray          ' Orders from the file
    Dim strDcFile As String             ' Daniel code file

    TimerStart "frmOnlineBroker.tmrGmaj"
    strDcFile = GetIniFileProperty("OutputFile", "", "GmajPro", AddSlash(App.Path) & "Provided\Provided.INI")
    If Len(strDcFile) = 0 Then
        tmrGmaj.Enabled = False
    Else
        If FileExist(AddSlash(App.Path) & strDcFile) Then
            tmrGmaj.Enabled = False
            
            Set astrOrders = New cGdArray
            astrOrders.FromFile AddSlash(App.Path) & strDcFile
            KillFile AddSlash(App.Path) & strDcFile
            
            If astrOrders.Size > 0 Then
                frmDanielConfirmation.ShowMe astrOrders, True
            End If
        
            frmMain.tbToolbar.Tools("ID_DanCodeWeb").Enabled = True
            frmMain.tbToolbar.Tools("ID_GmajPro").Enabled = True
        End If
    End If
    TimerEnd "frmOnlineBroker.tmrGmaj", tmrGmaj.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.tmrGmaj_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrHeartbeat_Timer
'' Description: Check the appropriate heartbeats for the brokers
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrHeartbeat_Timer()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim BrokerObj As cBroker            ' Broker object
    
    TimerStart "frmOnlineBroker.tmrHeartbeat"
    For lIndex = 1 To kNumBrokers - 1
        If g.Broker.IsLiveAccount(lIndex) Then
            If g.Broker.IsBrokerUser(lIndex) Then
                Select Case lIndex
'                    Case eTT_AccountType_CtgPfg:
'                        g.CtgPfg.CheckHeartbeat
'                    Case eTT_AccountType_FintecPfg:
'                        g.FintecPfg.CheckHeartbeat
'                    Case eTT_AccountType_LindWaldock:
'                        g.LindWaldock.CheckHeartbeat
'                    Case eTT_AccountType_ManExpress:
'                        g.ManExpress.CheckHeartbeat
'                    Case eTT_AccountType_PFG:
'                        g.PFG.CheckHeartbeat
                    Case eTT_AccountType_TransAct:
                        g.Transact.CheckHeartbeat
                    Case Else
                        Set BrokerObj = g.Broker.Broker(lIndex)
                        If Not BrokerObj Is Nothing Then
                            BrokerObj.CheckHeartbeat
                        End If
                        
                End Select
            End If
        End If
    Next lIndex
    
    ' 01/07/2012 DAJ: If the flag file exists, dump some trade console profiling ( checking
    ' it here because this timer only goes off every 5 seconds )...
    If FormIsLoaded("frmTTSummary") Then
        frmTTSummary.DumpProfile = FileExist(AddSlash(App.Path) & "TTProfile.FLG")
    End If
    TimerEnd "frmOnlineBroker.tmrHeartbeat", tmrHeartbeat.Interval
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.tmrHeartbeat_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrMessages_Timer
'' Description: Handle queued up messages
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrMessages_Timer()
On Error GoTo ErrSection:

    Dim msg As gdOCX.gdAppMailMsg       ' Message object received
    Dim strMessage As String            ' Message received
    Dim nBroker As eTT_AccountType      ' Broker for the message
    
    TimerStart "frmOnlineBroker.tmrMessages"
    m.bProcessingMessageQueue = True
    
    Do While m.Messages.Count > 0
        Set msg = m.Messages(1)
        m.Messages.Remove 1
        strMessage = msg.Message
        DumpDebug "Message removed from queue: " & LogStringForMessage(msg)
        
        nBroker = g.Broker.BrokerForControlName(msg.FromControlName)
        g.Broker.HandleMessage nBroker, msg.MsgType, strMessage
    Loop
    
    If m.Messages.Count = 0 Then
        MessageTimerEnabled = False
    End If
    TimerEnd "frmOnlineBroker.tmrMessages", tmrMessages.Interval
    
ErrExit:
    m.bProcessingMessageQueue = False
    Exit Sub
    
ErrSection:
    m.bProcessingMessageQueue = False
    RaiseError "frmOnlineBroker.tmrMessages_Timer"
    
End Sub

#If 0 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLastKnownCtgPfgSymbols
'' Description: Get the last known Capital Trading Group PFG symbols
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetLastKnownCtgPfgSymbols()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' List of symbols
    Dim astrLine As New cGdArray        ' Line in the file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lLastDate As Long               ' Last successful connection to broker
    
    If g.Broker.IsBrokerUser(eTT_AccountType_CtgPfg) Then
        lLastDate = CLng(Val(DecryptFromHex(GetIniFileProperty("Last", "", "Connect", g.CtgPfg.IniFile))))
        
        If g.Broker.IsBrokerSimUser(eTT_AccountType_CtgPfg) Or (lLastDate >= Date - 30) Then
            If astrSymbols.FromFile(AddSlash(App.Path) & "Provided\CtgPfgToGen.TXT") Then
                For lIndex = 0 To astrSymbols.Size - 1
                    astrLine.Clear
                    astrLine.SplitFields astrSymbols(lIndex), vbTab
                    
                    If Len(astrLine(6)) = 0 Then astrLine(6) = "1"
                    
                    If (Len(astrLine(1)) > 0) And (astrLine(6) = "1") Then
                        g.RealTime.AddBrokerRtSymbol astrLine(1) & "-", "CtgPfg/D"
                    End If
                Next lIndex
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.GetLastKnownCtgPfgSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLastKnownFintecPfgSymbols
'' Description: Get the last known Fintec Group PFG symbols
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetLastKnownFintecPfgSymbols()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' List of symbols
    Dim astrLine As New cGdArray        ' Line in the file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lLastDate As Long               ' Last successful connection to broker
    
    If g.Broker.IsBrokerUser(eTT_AccountType_FintecPfg) Then
        lLastDate = CLng(Val(DecryptFromHex(GetIniFileProperty("Last", "", "Connect", g.FintecPfg.IniFile))))
        
        If g.Broker.IsBrokerSimUser(eTT_AccountType_FintecPfg) Or (lLastDate >= Date - 30) Then
            If astrSymbols.FromFile(AddSlash(App.Path) & "Provided\FintecPfgToGen.TXT") Then
                For lIndex = 0 To astrSymbols.Size - 1
                    astrLine.Clear
                    astrLine.SplitFields astrSymbols(lIndex), vbTab
                    
                    If Len(astrLine(6)) = 0 Then astrLine(6) = "1"
                    
                    If (Len(astrLine(1)) > 0) And (astrLine(6) = "1") Then
                        g.RealTime.AddBrokerRtSymbol astrLine(1) & "-", "FintecPfg/D"
                    End If
                Next lIndex
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.GetLastKnownFintecPfgSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLastKnownLindWaldockSymbols
'' Description: Get the last known Lind Waldock symbols
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetLastKnownLindWaldockSymbols()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' List of symbols
    Dim astrLine As New cGdArray        ' Line in the file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lLastDate As Long               ' Last successful connection to Photon
    
    If g.Broker.IsBrokerUser(eTT_AccountType_LindWaldock) Then
        lLastDate = CLng(Val(DecryptFromHex(GetIniFileProperty("Last", "", "Connect", AddSlash(App.Path) & "LindWaldock.INI"))))
        
        If g.Broker.IsBrokerSimUser(eTT_AccountType_LindWaldock) Or (lLastDate >= Date - 30) Then
            If astrSymbols.FromFile(AddSlash(App.Path) & "Provided\LwToGen.TXT") Then
                For lIndex = 0 To astrSymbols.Size - 1
                    astrLine.Clear
                    astrLine.SplitFields astrSymbols(lIndex), vbTab
                    
                    If Len(astrLine(6)) = 0 Then astrLine(6) = "1"
                    
                    If (Len(astrLine(2)) > 0) And (astrLine(6) = "1") Then
                        g.RealTime.AddBrokerRtSymbol astrLine(2) & "-", "LindWaldock/D"
                    End If
                Next lIndex
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.GetLastKnownLindWaldockSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLastKnownManExpressSymbols
'' Description: Get the last known Man Express symbols
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetLastKnownManExpressSymbols()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' List of symbols
    Dim astrLine As New cGdArray        ' Line in the file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lLastDate As Long               ' Last successful connection to broker

    If g.Broker.IsBrokerUser(eTT_AccountType_ManExpress) Then
        lLastDate = CLng(Val(DecryptFromHex(GetIniFileProperty("Last", "", "Connect", AddSlash(App.Path) & "ManExpress.INI"))))

        If g.Broker.IsBrokerSimUser(eTT_AccountType_ManExpress) Or (lLastDate >= Date - 30) Then
            If astrSymbols.FromFile(AddSlash(App.Path) & "Provided\MxToGen.TXT") Then
                For lIndex = 0 To astrSymbols.Size - 1
                    astrLine.Clear
                    astrLine.SplitFields astrSymbols(lIndex), vbTab

                    If Len(astrLine(6)) = 0 Then astrLine(6) = "1"

                    If (Len(astrLine(2)) > 0) And (astrLine(6) = "1") Then
                        g.RealTime.AddBrokerRtSymbol astrLine(2) & "-", "ManExpress/D"
                    End If
                Next lIndex
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmOnlineBroker.GetLastKnownManExpressSymbols"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLastKnownPfgSymbols
'' Description: Get the last known PFG symbols
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetLastKnownPfgSymbols()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' List of symbols
    Dim astrLine As New cGdArray        ' Line in the file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lLastDate As Long               ' Last successful connection to broker
    
    If g.Broker.IsBrokerUser(eTT_AccountType_PFG) Then
        lLastDate = CLng(Val(DecryptFromHex(GetIniFileProperty("Last", "", "Connect", AddSlash(App.Path) & "Pfg.INI"))))
        
        If g.Broker.IsBrokerSimUser(eTT_AccountType_PFG) Or (lLastDate >= Date - 30) Then
            If astrSymbols.FromFile(AddSlash(App.Path) & "Provided\PfgToGen2.TXT") Then
                For lIndex = 0 To astrSymbols.Size - 1
                    astrLine.Clear
                    astrLine.SplitFields astrSymbols(lIndex), vbTab
                    
                    If Len(astrLine(6)) = 0 Then astrLine(6) = "1"
                    
                    If (Len(astrLine(1)) > 0) And (astrLine(6) = "1") Then
                        g.RealTime.AddBrokerRtSymbol astrLine(1) & "-", "PFG/D"
                    End If
                Next lIndex
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.GetLastKnownPfgSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLastKnownTransactSymbols
'' Description: Get the last known TransAct symbols
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetLastKnownTransactSymbols()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' List of symbols
    Dim astrFile As New cGdArray        ' Symbol translation file
    Dim astrLine As New cGdArray        ' Line out of the symbol translation file
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSymbols As String            ' Default symbols
    Dim strKey As String                ' Key into the registry
    
    strKey = "Software\Genesis Financial Data Services\Navigator Suite\General"
    
    If FileExist(AddSlash(App.Path) & "Provided\LKTran.SYM") Then
        astrSymbols.FromFile (AddSlash(App.Path) & "Provided\LKTran.SYM")
        astrSymbols.Serialize AddSlash(App.Path) & "Provided\LKSyms.TRN", True
        SetRegistryValue rkLocalMachine, strKey, "LKTSC", CalcFileCrc(AddSlash(App.Path) & "Provided\LKSyms.TRN"), True
        KillFile AddSlash(App.Path) & "Provided\LKTran.SYM"
    End If
    
    If CalcFileCrc(AddSlash(App.Path) & "Provided\LKSyms.TRN") <> GetRegistryValue(rkLocalMachine, strKey, "LKTSC", 0&) Then
        KillFile AddSlash(App.Path) & "Provided\LKSyms.TRN"
    End If
    
    If astrSymbols.Serialize(AddSlash(App.Path) & "Provided\LKSyms.TRN", False) = False Then
        ' DAJ 04/22/2014: Because of the new CME rules, TransAct is going to be required to
        ' get rid of the "simuser" demo account ( they will need to go to individual simulated
        ' accounts per user ).  Because of this, we are going to stop connecting to "simuser",
        ' but utilize the BRKRDEMO enablement to still give the user real-time data on the same
        ' handful of symbols for the first month ( or until they login with a real login )...
        'If (g.Broker.IsBrokerSimUser(eTT_AccountType_TransAct) = True) And (g.Transact.UserName = g.Transact.SimUserUserName) Then
        If g.Broker.IsBrokerSimUser(eTT_AccountType_TransAct) = True Then
            'strSymbols = "G6A-,G6B-,G6C-,G6E-,G6J-,G6S-,E7-,EMD-,ER2-,ES-,GE-,NQ-,QG-,QM-,XK-,XY-,YM-,ZB-,ZC-,ZF-,ZG-,ZI-,ZL-,ZN-,ZS-,ZT-,ZU-,ZW-"
            strSymbols = "G6E-,ER2-,ES-,NQ-,YM-,ZB-,ZN-"
            astrSymbols.SplitFields strSymbols, ","
            astrSymbols.Serialize AddSlash(App.Path) & "Provided\LKSyms.TRN", True
            SetRegistryValue rkLocalMachine, strKey, "LKTSC", CalcFileCrc(AddSlash(App.Path) & "Provided\LKSyms.TRN"), True
        End If
    End If
        
    If astrSymbols.Size > 0 Then
        astrSymbols.Sort
        If astrFile.FromFile(AddSlash(App.Path) & "Provided\TrnToGen.TXT") Then
            For lIndex = 0 To astrFile.Size - 1
                astrLine.Clear
                astrLine.SplitFields astrFile(lIndex), vbTab
                
                If Len(astrLine(6)) = 0 Then astrLine(6) = "1"
                
                If astrSymbols.BinarySearch(astrLine(2) & "-") And (astrLine(6) = "1") Then
                    g.RealTime.AddBrokerRtSymbol astrLine(2) & "-", "Transact/D"
                End If
            Next lIndex
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.GetLastKnownTransactSymbols"
    
End Sub
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLastKnownBrokerSymbols
'' Description: Get the last known symbols for all appropriate brokerages
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetLastKnownBrokerSymbols()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim BrokerObj As cBroker            ' Broker object
    Dim strSmizx As String              ' Override list for E_SMIZX
    Dim astrSmiZx As cGdArray           ' Override list for E_SMIZX broken into an array
    
    For lIndex = 1 To kNumBrokers - 1
        If g.Broker.IsLiveAccount(lIndex) Then
            If g.Broker.IsBrokerUser(lIndex) Then
                Select Case lIndex
'                    Case eTT_AccountType_CtgPfg:
'                        GetLastKnownCtgPfgSymbols
'                    Case eTT_AccountType_FintecPfg:
'                        GetLastKnownFintecPfgSymbols
'                    Case eTT_AccountType_LindWaldock:
'                        GetLastKnownLindWaldockSymbols
'                    Case eTT_AccountType_ManExpress:
'                        GetLastKnownManExpressSymbols
'                    Case eTT_AccountType_PFG:
'                        GetLastKnownPfgSymbols
                    Case eTT_AccountType_TransAct:
                        If Not g.Transact Is Nothing Then
                            g.Transact.AddBrokerRtSymbols
                        End If
                    Case Else
                        Set BrokerObj = g.Broker.Broker(lIndex)
                        If Not BrokerObj Is Nothing Then
                            BrokerObj.AddBrokerRtSymbols
                        End If
                        
                End Select
            End If
        End If
    Next lIndex

    ' DAJ 05/10/2013: The Elliott Wave folks want the YC real-time, but don't connect
    ' to a broker.  Since the SFE symbols can only be streamed via a broker override,
    ' we will use the enablement here to override them...
    If HasModule("E_SMIZX") Then
        strSmizx = DecryptFromHex(GetProvidedProperty("SmiZx"))
        If Len(strSmizx) > 0 Then
            Set astrSmiZx = New cGdArray
            astrSmiZx.SplitFields strSmizx, ","
            
            For lIndex = 0 To astrSmiZx.Size - 1
                g.RealTime.AddBrokerRtSymbol astrSmiZx(lIndex), "SmiZx/D"
            Next lIndex
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.GetLastKnownBrokerSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Send a string to the log file for the day
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DumpDebug(ByVal strMessage As String)
On Error Resume Next

#If 0 Then

    Dim fh As Integer                   ' File handle to open file with

    fh = FreeFile
    Open AddSlash(App.Path) & "Brokers\TN" & Format(Now, "YYYYMMDD") & ".LOG" For Append Shared As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " (" & Str(gdTickCount) & ") - " & strMessage
        Close #fh
    End If

#Else

    Static LogFile As cLogFile
    If LogFile Is Nothing Then
        Set LogFile = New cLogFile
        LogFile.OpenFile AddSlash(App.Path) & "Brokers\TN*.LOG"
    End If
    LogFile.WriteText strMessage

#End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddMessageToQueue
'' Description: Add the given message to the message queue
'' Inputs:      Message, Key
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddMessageToQueue(ByVal msg As gdOCX.gdAppMailMsg, Optional ByVal strKey As String = "")
On Error GoTo ErrSection:

    If Len(strKey) = 0 Then
        DumpDebug "Message added to queue: " & LogStringForMessage(msg)
        m.Messages.Add msg
    Else
        If m.Messages.Exists(strKey) Then
            DumpDebug "Message replaced in queue: " & LogStringForMessage(msg)
            m.Messages(strKey) = msg
        Else
            DumpDebug "Message added to queue: " & LogStringForMessage(msg)
            m.Messages.Add msg, strKey
        End If
    End If
    
    MessageTimerEnabled = True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmOnlineBroker.AddMessageToQueue"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LogStringForMessage
'' Description: Return a log string for the given app mail message
'' Inputs:      App Mail Message
'' Returns:     Log String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LogStringForMessage(ByVal msg As gdOCX.gdAppMailMsg) As String
On Error GoTo ErrSection:

    LogStringForMessage = "Number = '" & Str(msg.MsgNumber) & "' ; ControlName = '" & msg.FromControlName & "' ; Type = '" & Str(msg.MsgType) & "' ; Message = '" & msg.Message & "'"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOnlineBroker.LogStringForMessage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrUpdateCharts_Timer
'' Description: Do an update visible charts for symbol ID's that are in the tag
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrUpdateCharts_Timer()
On Error GoTo ErrSection:

    Dim astrSymbolIDs As New cGdArray   ' Array of symbol ID's to refresh
    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long

    TimerStart "frmOnlineBroker.tmrUpdateCharts"
    tmrUpdateCharts.Enabled = False
    astrSymbolIDs.SplitFields tmrUpdateCharts.Tag, ","
    tmrUpdateCharts.Tag = ""
    
    For lIndex = 0 To astrSymbolIDs.Size - 1
        If Len(astrSymbolIDs(lIndex)) > 0 Then
            lSymbolID = Val(astrSymbolIDs(lIndex))
            If lSymbolID > 0 Then
                UpdateVisibleCharts eRedo1_Scrolled, lSymbolID, , True      '5841
                CheckFillsToJournal lSymbolID
            End If

'JM 11-02-2010: don't think this is needed since the call to UpdateVisibleCharts above will also take care of synthetics
            ' also update any synthetics (e.g. ES1 charts)
'            lSymbolID = GetSymbolID(ConvertSynthetic(GetSymbol(lSymbolID), True))
'            If lSymbolID > 0 Then
'                UpdateVisibleCharts eRedo1_Scrolled, lSymbolID
'            End If

        End If
    Next lIndex
    TimerEnd "frmOnlineBroker.tmrUpdateCharts", tmrUpdateCharts.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.tmrUpdateCharts_Timer"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtSalmonCallback_Change
'' Description: Handle a message back from the Salmon client
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSalmonCallback_Change()
    
    Dim strText As String               ' Text back from the window
    
    ' this text window is just being used for the Salmon DLL callback functionality
    strText = Trim(txtSalmonCallback.Text)
    If Len(strText) > 0 And Not g.RealTime Is Nothing Then
        g.RealTime.SalmonCallback strText
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckFillsToJournal
'' Description: Check the fills to journal collection for the given symbol ID
'' Inputs:      Symbol ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckFillsToJournal(ByVal lSymbolID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object
    Dim alToDelete As cGdArray          ' Array of indexes to delete
    
    Set alToDelete = New cGdArray
    alToDelete.Create eGDARRAY_Longs
    
    For lIndex = 1 To m.FillsToJournal.Count
        Set Fill = m.FillsToJournal(lIndex)
        If Fill.SymbolID = lSymbolID Then
            g.TnJournal.AutoJournalForFill Fill
            alToDelete.Add lIndex
        End If
    Next lIndex
    
    For lIndex = 0 To alToDelete.Size - 1
        m.FillsToJournal.Remove alToDelete(lIndex)
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOnlineBroker.CheckFillsToJournal"
    
End Sub

