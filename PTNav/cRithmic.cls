VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRithmic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cRithmic.cls
'' Description: Wrapper class for interacting with the Rithmic server
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 09/15/2010   DAJ         Added Option Navigator calls
'' 10/12/2010   DAJ         Fix for "Not connected to account" message
'' 10/26/2010   DAJ         Logging and efficiency enhancements
'' 10/29/2010   DAJ         Fix for carried position mismatch fix on wrong side
'' 11/01/2010   DAJ         Added Optimus, OpVest, and Vision (Rithmic Brokers)
'' 11/09/2010   DAJ         Changed the name of the connect file for Optimus
'' 11/18/2010   DAJ         Added Starting Genesis ID for Option Navigator
'' 11/29/2010   DAJ         Added code to convert NYBOT futures
'' 11/30/2010   DAJ         Added COMEX to symbol conversion routines
'' 12/03/2010   DAJ         Set last connected date in INI file for broker overrides
'' 12/10/2010   DAJ         Added Zen-Fire
'' 01/26/2011   DAJ         Added capabilities for stocks and stock options
'' 03/03/2011   DAJ         Added NLX for symbol conversion
'' 03/07/2011   DAJ         Moved over to broker base class
'' 04/05/2011   DAJ         Added SymbolInformation call, added SellShort & SellShortExempt
'' 05/27/2011   DAJ         Send EntryType (Manual/Automated) along with an order
'' 10/20/2011   DAJ         Clear "BeginSent" flag on End
'' 12/12/2011   DAJ         Create Position message for flat position w/ carried fill
'' 02/01/2012   DAJ         Fixed "Object variable..." error in SendFlatPositions
'' 05/22/2012   DAJ         Fixed the brokers area path for ZenFire
'' 07/16/2012   DAJ         ZanerRithmic, ZanerZenFire
'' 07/17/2012   DAJ         AlpariZenFire
'' 07/24/2012   DAJ         Changed the Rithmic login stuff
'' 08/28/2012   DAJ         CarriedFillFromInfo, First Field Begin/End
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 06/27/2013   DAJ         Added 'EDGX' exchange when converting from a Rithmic symbol to Genesis
'' 07/16/2013   DAJ         Trade Routes for Rithmic, BATS stocks convert to @BATS
'' 08/20/2014   DAJ         Added 'ICE' exchange when converting symbols
'' 10/23/2014   DAJ         Added symbol conversion for calendar spreads on futures
'' 11/14/2014   DAJ         Added support for MIT orders
'' 01/29/2015   DAJ         Added ContractInfo calls; Pass credentials twice
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Broker As cBroker                   ' Broker base class object
    Symbols As cBrokerSymbols           ' Broker symbol conversion object

    strUserID As String                 ' User ID to logon to the server with
    strPassword As String               ' Password to logon to the server with
    strCnctFile As String               ' INI file with connection information
    
    OrdersRequestedFor As cGdTree       ' Collection of accounts that orders have been requested for
    bOrdersBeginSent As Boolean         ' A Begin has been sent for orders
    FillsRequestedFor As cGdTree        ' Collection of accounts that fills have been requested for
    bFillsBeginSent As Boolean          ' A Begin has been sent for fills
    PositionsRequestedFor As cGdTree    ' Collection of accounts that positions have been requested for
    bPositionsBeginSent As Boolean      ' A Begin has been sent for positions
    
    astrExchanges As cGdArray           ' Exchanges the user is entitled for through Rithmic
    astrIbs As cGdArray                 ' Array of FCM/IB combinations
    IbIdsByAccount As cGdTree           ' Collection of FCM/IB by account
    TradeRoutes As cGdTree              ' Collection of trade routes
    
    CarPosFills As cGdTree              ' Collection of carried position information
End Type
Private m As mPrivate

Public Property Get Broker() As cBroker
    Set Broker = m.Broker
End Property

Public Property Get UserID() As String
    UserID = m.strUserID
End Property
Public Property Let UserID(ByVal strUserID As String)
    m.strUserID = strUserID
End Property

Public Property Get Password() As String
    Password = m.strPassword
End Property
Public Property Let Password(ByVal strPassword As String)
    m.strPassword = strPassword
    m.Broker.AskedPassword = True
End Property

Private Property Get ConnectString() As String
    ConnectString = m.strUserID & vbTab & m.strPassword & vbTab & "Trade Navigator" & vbTab & FormatVersion & "." & App.Revision & vbTab & m.strCnctFile & vbTab & m.strUserID & vbTab & m.strPassword
End Property
Private Property Get ConnectStringLog() As String
    ConnectStringLog = m.strUserID & vbTab & SetStrToChar(m.strPassword, "*") & vbTab & "Trade Navigator" & vbTab & FormatVersion & "." & App.Revision & vbTab & m.strCnctFile & vbTab & m.strUserID & vbTab & SetStrToChar(m.strPassword, "*")
End Property

Public Property Get EnabledSymbols() As cGdArray
    Set EnabledSymbols = Nothing
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the object based on what broker it is
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim strSymbolFile As String         ' Symbol conversion file
    Dim astrSymbols As New cGdArray     ' Array of symbols from a file
    
    Select Case nBroker
        Case eTT_AccountType_AlpariZenFire
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "AlpariZenFire"
            m.Broker.IniFile = AddSlash(App.Path) & "AlpariZenFire.INI"
            m.Broker.BrokerName = "Alpari (Zen-Fire)"
            m.Broker.ProcessName = "GenAlpariZenFire"
            m.Broker.ProcessPath = AddSlash(App.Path) & "..\Brokers\AlpariZenFire\GenAlpariZenFire.EXE"
            m.Broker.ControlID = "AlpariZenFire"
            strSymbolFile = AddSlash(App.Path) & "Provided\AlpariZenToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\AlpariZenCnct.INI"
            
        Case eTT_AccountType_Optimus
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "Optimus"
            m.Broker.IniFile = AddSlash(App.Path) & "Optimus.INI"
            m.Broker.BrokerName = "Optimus"
            m.Broker.ProcessName = "GenOptimus"
            m.Broker.ProcessPath = AddSlash(App.Path) & "..\Brokers\Optimus\GenOptimus.EXE"
            m.Broker.ControlID = "Optimus"
            strSymbolFile = AddSlash(App.Path) & "Provided\OmsToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\OmsCnct.INI"
            
        Case eTT_AccountType_OpVest
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "OpVest"
            m.Broker.IniFile = AddSlash(App.Path) & "OpVest.INI"
            m.Broker.BrokerName = "OpVest"
            m.Broker.ProcessName = "GenOpVest"
            m.Broker.ProcessPath = AddSlash(App.Path) & "..\Brokers\OpVest\GenOpVest.EXE"
            m.Broker.ControlID = "OpVest"
            strSymbolFile = AddSlash(App.Path) & "Provided\OpvToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\OpvCnct.INI"
            
        Case eTT_AccountType_Rithmic
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "Rithmic"
            m.Broker.IniFile = AddSlash(App.Path) & "Rithmic.INI"
            m.Broker.BrokerName = "Rithmic"
            m.Broker.ProcessName = "GenRithmic"
            m.Broker.ProcessPath = AddSlash(App.Path) & "..\Brokers\Rithmic\GenRithmic.EXE"
            m.Broker.ControlID = "Rithmic"
            strSymbolFile = AddSlash(App.Path) & "Provided\RthToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\RithCnct.INI"
            
        Case eTT_AccountType_Vision
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "Vision"
            m.Broker.IniFile = AddSlash(App.Path) & "Vision.INI"
            m.Broker.BrokerName = "Vision"
            m.Broker.ProcessName = "GenVision"
            m.Broker.ProcessPath = AddSlash(App.Path) & "..\Brokers\Vision\GenVision.EXE"
            m.Broker.ControlID = "Vision"
            strSymbolFile = AddSlash(App.Path) & "Provided\VsnToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\VsnCnct.INI"
            
        Case eTT_AccountType_ZanerRithmic
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "ZanerRithmic"
            m.Broker.IniFile = AddSlash(App.Path) & "ZanerRithmic.INI"
            m.Broker.BrokerName = "Zaner (Rithmic)"
            m.Broker.ProcessName = "GenZanerRithmic"
            m.Broker.ProcessPath = AddSlash(App.Path) & "..\Brokers\ZanerRithmic\GenZanerRithmic.EXE"
            m.Broker.ControlID = "ZanerRithmic"
            strSymbolFile = AddSlash(App.Path) & "Provided\ZanerRthToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\ZanerRthCnct.INI"
            
        Case eTT_AccountType_ZanerZenFire
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "ZanerZenFire"
            m.Broker.IniFile = AddSlash(App.Path) & "ZanerZenFire.INI"
            m.Broker.BrokerName = "Zaner (Zen-Fire)"
            m.Broker.ProcessName = "GenZanerZenFire"
            m.Broker.ProcessPath = AddSlash(App.Path) & "..\Brokers\ZanerZenFire\GenZanerZenFire.EXE"
            m.Broker.ControlID = "ZanerZenFire"
            strSymbolFile = AddSlash(App.Path) & "Provided\ZanerZenToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\ZanerZenCnct.INI"
            
        Case eTT_AccountType_ZenFire
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "ZenFire"
            m.Broker.IniFile = AddSlash(App.Path) & "ZenFire.INI"
            m.Broker.BrokerName = "Zen-Fire"
            m.Broker.ProcessName = "GenZenFire"
            m.Broker.ProcessPath = AddSlash(App.Path) & "..\Brokers\ZenFire\GenZenFire.EXE"
            m.Broker.ControlID = "ZenFire"
            strSymbolFile = AddSlash(App.Path) & "Provided\ZenToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\ZenCnct.INI"
            
    End Select
    
    m.Broker.CopyProcess AddSlash(App.Path) & "..\Brokers\GenRithmic.EXE"
            
    Set m.Symbols = New cBrokerSymbols
    m.Symbols.FromFile strSymbolFile
    
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Connect) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Disconnect) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_SwitchAccounts) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_SwitchAccountsMode) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ConnectInfo) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ChangePassword) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Refresh) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewActivity) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_BrokerView) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewOnline) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_VerifyPositions) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_AccountDetails) = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.Init"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowConfig
'' Description: Determines whether we should show the config form on connect
'' Inputs:      None
'' Returns:     True if show configuration form, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowConfig() As Boolean
On Error GoTo ErrSection:

    ShowConfig = (Len(m.strUserID) = 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.ShowConfig"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowBrokerConnectionInfo
'' Description: Allow the user to view their broker connection information
'' Inputs:      New?, User Name
'' Returns:     True if dialog OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowBrokerConnectionInfo(Optional ByVal bNew As Boolean = False, Optional ByVal strUserName As String = "") As Boolean
On Error GoTo ErrSection:

    ShowBrokerConnectionInfo = GetLoginInfo(strUserName, True)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.ShowBrokerConnectionInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowLogin
'' Description: Show the user the login form
'' Inputs:      New User Name, Are we switching?
'' Returns:     True if connect, False othwerwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowLogin(Optional ByVal strNewUserName As String = "", Optional ByVal bSwitching As Boolean = False)
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lTimeOut As Long                ' Timeout variable

    If (m.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected) Then
        If (m.Broker.AskedPassword = True) And ((strNewUserName = m.strUserID) Or (Len(strNewUserName) = 0)) Then
            bReturn = True
        Else
            bReturn = GetLoginInfo(strNewUserName)
            m.Broker.AskedPassword = bReturn
        End If
    Else
        If ((strNewUserName <> m.strUserID) And (Len(strNewUserName) > 0)) Or (bSwitching = True) Then
            bReturn = GetLoginInfo(strNewUserName, bSwitching)
            If bReturn = True Then
                m.Broker.AskedPassword = True
                m.Broker.Disconnect False, "Changing user names to '" & strNewUserName & "'"
                
                lTimeOut = 0&
                Do While (m.Broker.ConnectionStatus <> eGDConnectionStatus_Disconnected) And (lTimeOut < 30&)
                    Sleep 1
                    lTimeOut = lTimeOut + 1&
                Loop
                
                bReturn = (m.Broker.ConnectionStatus = eGDConnectionStatus_Disconnected)
            End If
        End If
    End If
    
    ShowLogin = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.ShowLogin"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendConnect
'' Description: Send a connect to the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendConnect()
On Error GoTo ErrSection:

    m.Broker.DumpDebug "Connecting to " & m.Broker.BrokerName & " as '" & m.strUserID & "'"
    SendRithmicMessage eGDRithmicMessageType_Connect, ConnectString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.SendConnect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendDisconnect
'' Description: Send a disconnect to the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendDisconnect()
On Error GoTo ErrSection:

    SendRithmicMessage eGDRithmicMessageType_Disconnect, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.SendDisconnect"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UnloadApp
'' Description: Unload the stand-alone application
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UnloadApp()
On Error GoTo ErrSection:

    SendRithmicMessage eGDRithmicMessageType_UnloadApp, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.UnloadApp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerSymbol
'' Description: Attempt to convert the Genesis symbol to the broker symbology
'' Inputs:      Genesis Symbol
'' Returns:     Broker Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerSymbol(ByVal strGenesisSymbol As String, Optional BrokerSym As cBrokerSymbol) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strGenesisBase As String        ' Base symbol for the Genesis symbol
    Dim strContract As String           ' Contract for the Genesis symbol
    Dim strYear As String               ' Contract year
    Dim lMonth As Long                  ' Contract month
    Dim OptSym As New cOptionSymbol     ' Option symbol object
    Dim astrSymbol As cGdArray          ' Symbol broken out into pieces
    Dim astrComponents As cGdArray      ' Spread components
    
    strReturn = ""
    If Len(strGenesisSymbol) > 0 Then
        Select Case UCase(SecurityType(strGenesisSymbol, True))
            Case "F"
                Set astrSymbol = New cGdArray
                astrSymbol.SplitFields strGenesisSymbol, "-"
                
                If astrSymbol.Size = 2 Then
                    strGenesisBase = astrSymbol(0)
                    Set BrokerSym = m.Symbols.GetByGenesisSymbol(strGenesisBase)
                    If Not BrokerSym Is Nothing Then
                        strContract = astrSymbol(1)
                        strYear = Left(strContract, 4)
                        lMonth = CLng(Val(Right(strContract, 2)))
                        
                        Select Case UCase(BrokerSym.BrokerExchange)
                            Case "CME", "CBOT", "NYMEX", "COMEX":
                                ' ESZ4
                                strReturn = BrokerSym.BrokerBase & MonthToCode(lMonth) & Right(strYear, 1)
                            Case "EUREX", "NLX":
                                ' FDAX0315
                                strReturn = BrokerSym.BrokerBase & Format(lMonth, "00") & Right(strYear, 2)
                            Case "NYBOT", "ICE":
                                ' KC FMZ0014!
                                strReturn = BrokerSym.BrokerBase & "  FM" & MonthToCode(lMonth) & "00" & Right(strYear, 2) & "!"
                        End Select
                    End If
                
                ElseIf astrSymbol.Size = 3 Then
                    strGenesisBase = astrSymbol(0)
                    Set BrokerSym = m.Symbols.GetByGenesisSymbol("S:" & strGenesisBase)
                    If Not BrokerSym Is Nothing Then
                        Set astrComponents = mDataNav.SpreadComponentsForSymbol(strGenesisSymbol)
                        If astrComponents.Size = 2 Then
                            astrComponents(0) = Replace(BrokerSymbol(astrComponents(0)), "!", "")
                            astrComponents(1) = Replace(BrokerSymbol(astrComponents(1)), "!", "")
                            
                            Select Case UCase(BrokerSym.BrokerExchange)
                                Case "CME", "CBOT", "NYMEX", "COMEX":
                                    ' ESZ4-ESH5
                                    strReturn = astrComponents(0) & "-" & astrComponents(1)
                                Case "EUREX", "NLX":
                                    ' FDAX1214-FDAX0315
                                    strReturn = astrComponents(0) & "-" & astrComponents(1)
                                Case "NYBOT", "ICE":
                                    ' KC FMZ0014-KC FMH0015
                                    strReturn = astrComponents(0) & "-" & astrComponents(1)
                            End Select
                        End If
                    End If
                End If
                
            Case "FO"
                OptSym.FromGenesis strGenesisSymbol
                Set BrokerSym = m.Symbols.GetByGenesisSymbol("O:" & OptSym.BaseSymbol)
                If Not BrokerSym Is Nothing Then
                    strReturn = OptSym.ToRithmic(Parse(BrokerSym.BrokerBase, ":", 2), BrokerSym.BrokerExchange, BrokerSym.StrikeMult, BrokerSym.StrikePriceFormat)
                End If
                
            Case "S"
                Set BrokerSym = m.Symbols.GetByGenesisSymbol("!")
                If Not BrokerSym Is Nothing Then
                    If InStr(strGenesisSymbol, "_P") <> 0 Then
                        strReturn = Replace(strGenesisSymbol, "_P", "-")
                    ElseIf InStr(strGenesisSymbol, "_W") <> 0 Then
                        strReturn = Replace(strGenesisSymbol, "_W", "+")
                    ElseIf InStr(strGenesisSymbol, "_R") <> 0 Then
                        strReturn = Replace(strGenesisSymbol, "_R", "^")
                    ElseIf InStr(strGenesisSymbol, "_U") <> 0 Then
                        strReturn = Replace(strGenesisSymbol, "_U", "=")
                    ElseIf InStr(strGenesisSymbol, "_") = 0 Then
                        strReturn = strGenesisSymbol
                    End If
                    
                    If InStr(strReturn, "@BATS") <> 0 Then
                        strReturn = Replace(strReturn, "@BATS", "")
                        BrokerSym.BrokerExchange = "BATS"
                    End If
                End If
                
            Case "SO"
                OptSym.FromGenesis strGenesisSymbol
                Set BrokerSym = m.Symbols.GetByGenesisSymbol("@")
                If Not BrokerSym Is Nothing Then
                    strReturn = OptSym.ToRithmic(OptSym.BaseSymbol, BrokerSym.BrokerExchange, BrokerSym.StrikeMult, BrokerSym.StrikePriceFormat)
                End If
                
        End Select
    End If
    
    BrokerSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.BrokerSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Convert a Rithmic symbol to a Genesis symbol
'' Inputs:      Broker Symbol, Broker Exchange, Broker Symbol object, For Position?
'' Returns:     Genesis Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbol(ByVal strBrokerSymbol As String, ByVal strBrokerExchange As String, Optional BrokerSym As cBrokerSymbol, Optional ByVal bForPosition As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strSymbol As String             ' Temporary symbol variable
    Dim strBrokerBase As String         ' Broker base symbol
    Dim strBrokerBase2 As String        ' Broker base symbol
    Dim lYear As Long                   ' Contract year
    Dim lMonth As Long                  ' Contract month
    Dim OptSym As New cOptionSymbol     ' Option symbol object
    Dim strOptPart As String            ' Option part of the symbol
    Dim astrComponents As cGdArray      ' Components of a spread symbol
    
    strReturn = ""
    Select Case UCase(strBrokerExchange)
        Case "CME", "CBOT", "NYMEX", "COMEX":
            If InStr(strBrokerSymbol, "-") > 0 Then
                ' Spread symbol: HEZ4-HEG5
                Set astrComponents = New cGdArray
                astrComponents.SplitFields strBrokerSymbol, "-"
                
                If astrComponents.Size = 2 Then
                    strSymbol = astrComponents(0)
                    strBrokerBase = Left(strSymbol, Len(strSymbol) - 2)
                    strSymbol = astrComponents(1)
                    strBrokerBase2 = Left(strSymbol, Len(strSymbol) - 2)
                    
                    If strBrokerBase = strBrokerBase2 Then
                        If bForPosition Then
                            Set BrokerSym = m.Symbols.GetByBrokerPositionSymbol("S:" & strBrokerBase)
                        Else
                            Set BrokerSym = m.Symbols.GetByBrokerSymbol("S:" & strBrokerBase)
                        End If
                        
                        If Not BrokerSym Is Nothing Then
                            astrComponents(0) = GenesisSymbol(astrComponents(0), strBrokerExchange, , bForPosition)
                            astrComponents(1) = GenesisSymbol(astrComponents(1), strBrokerExchange, , bForPosition)
                            strReturn = mDataNav.SpreadSymbolForComponents(astrComponents)
                        End If
                    End If
                End If
                
            ElseIf InStr(strBrokerSymbol, ":") = 0 Then
                ' Futures symbol: HEZ4
                strSymbol = Parse(strBrokerSymbol, " ", 1)
                If Len(strSymbol) > 2 Then
                    strBrokerBase = Left(strSymbol, Len(strSymbol) - 2)
                    If InStr(strBrokerSymbol, " ") > 0 Then
                        strBrokerBase = "O:" & strBrokerBase
                    End If
                    If bForPosition Then
                        Set BrokerSym = m.Symbols.GetByBrokerPositionSymbol(strBrokerBase)
                    Else
                        Set BrokerSym = m.Symbols.GetByBrokerSymbol(strBrokerBase)
                    End If
                    
                    If Not BrokerSym Is Nothing Then
                        lYear = CLng(Val(Right(strSymbol, 1))) + 2000
                        lMonth = CodeToMonth(Mid(strSymbol, Len(strSymbol) - 1, 1))
                        Do While (lYear * 100 + lMonth) < (Year(Date) * 100 + Month(Date))
                            lYear = lYear + 10
                        Loop
                    
                        strReturn = BrokerSym.GenesisBase & "-" & Str(lYear) & Format(lMonth, "00")
                    
                        If InStr(strBrokerSymbol, " ") > 0 Then
                            strOptPart = Parse(strBrokerSymbol, " ", 2)
                            
                            OptSym.BaseSymbol = Parse(strReturn, ":", 2)
                            OptSym.IsPut = (UCase(Left(strOptPart, 1)) = "P")
                            OptSym.Strike = Val(Mid(strOptPart, 2)) * BrokerSym.StrikeMult
                            
                            strReturn = OptSym.ToGenesis
                        End If
                    End If
                End If
            End If
            
        Case "EUREX", "NLX":
            If InStr(strBrokerSymbol, "-") > 0 Then
                ' Spread symbol: FDAX1214-FDAX0315
                Set astrComponents = New cGdArray
                astrComponents.SplitFields strBrokerSymbol, "-"
                
                If astrComponents.Size = 2 Then
                    strSymbol = astrComponents(0)
                    strBrokerBase = Left(strSymbol, Len(strSymbol) - 4)
                    strSymbol = astrComponents(1)
                    strBrokerBase2 = Left(strSymbol, Len(strSymbol) - 4)
                    
                    If strBrokerBase = strBrokerBase2 Then
                        If bForPosition Then
                            Set BrokerSym = m.Symbols.GetByBrokerPositionSymbol("S:" & strBrokerBase)
                        Else
                            Set BrokerSym = m.Symbols.GetByBrokerSymbol("S:" & strBrokerBase)
                        End If
                        
                        If Not BrokerSym Is Nothing Then
                            astrComponents(0) = GenesisSymbol(astrComponents(0), strBrokerExchange, , bForPosition)
                            astrComponents(1) = GenesisSymbol(astrComponents(1), strBrokerExchange, , bForPosition)
                            strReturn = mDataNav.SpreadSymbolForComponents(astrComponents)
                        End If
                    End If
                End If
            
            ElseIf Len(strBrokerSymbol) > 4 Then
                ' Futures symbol: FDAX0315
                strBrokerBase = Left(strBrokerSymbol, Len(strBrokerSymbol) - 4)
                If bForPosition Then
                    Set BrokerSym = m.Symbols.GetByBrokerPositionSymbol(strBrokerBase)
                Else
                    Set BrokerSym = m.Symbols.GetByBrokerSymbol(strBrokerBase)
                End If
                    
                If Not BrokerSym Is Nothing Then
                    strReturn = BrokerSym.GenesisBase & "-20" & Right(strBrokerSymbol, 2) & Mid(strBrokerSymbol, Len(strBrokerSymbol) - 4 + 1, 2)
                End If
            End If
            
        Case "NYBOT", "ICE":
            If InStr(strBrokerSymbol, "-") > 0 Then
                ' Spread symbol: KC FMZ0014-KC FMH0015
                Set astrComponents = New cGdArray
                astrComponents.SplitFields strBrokerSymbol, "-"
                
                If astrComponents.Size = 2 Then
                    strSymbol = astrComponents(0)
                    strBrokerBase = Parse(strSymbol, " ", 1)
                    strSymbol = astrComponents(1)
                    strBrokerBase2 = Parse(strSymbol, " ", 1)
                    
                    If strBrokerBase = strBrokerBase2 Then
                        If bForPosition Then
                            Set BrokerSym = m.Symbols.GetByBrokerPositionSymbol("S:" & strBrokerBase)
                        Else
                            Set BrokerSym = m.Symbols.GetByBrokerSymbol("S:" & strBrokerBase)
                        End If
                        
                        If Not BrokerSym Is Nothing Then
                            astrComponents(0) = GenesisSymbol(astrComponents(0) & "!", strBrokerExchange, , bForPosition)
                            astrComponents(1) = GenesisSymbol(astrComponents(1) & "!", strBrokerExchange, , bForPosition)
                            strReturn = mDataNav.SpreadSymbolForComponents(astrComponents)
                        End If
                    End If
                End If
            
            ElseIf (InStr(strBrokerSymbol, " ") <> 0) And (Right(strBrokerSymbol, 1) = "!") Then
                ' Futures symbol: KC FMZ0014!
                strBrokerBase = Parse(strBrokerSymbol, " ", 1)
                If bForPosition Then
                    Set BrokerSym = m.Symbols.GetByBrokerPositionSymbol(strBrokerBase)
                Else
                    Set BrokerSym = m.Symbols.GetByBrokerSymbol(strBrokerBase)
                End If
                    
                If Not BrokerSym Is Nothing Then
                    strReturn = BrokerSym.GenesisBase & "-20" & Mid(strBrokerSymbol, Len(strBrokerSymbol) - 2, 2) & Format(CodeToMonth(Mid(strBrokerSymbol, Len(strBrokerSymbol) - 5, 1)), "00")
                End If
            End If
            
        Case "BATS", "EDGX"
            Set BrokerSym = m.Symbols.GetByBrokerSymbol("!")
            If Not BrokerSym Is Nothing Then
                If InStr(strBrokerSymbol, ".") <> 0 Then
                    strReturn = Replace(strBrokerSymbol, ".", "_")
                ElseIf InStr(strBrokerSymbol, "-") <> 0 Then
                    strReturn = Replace(strBrokerSymbol, "-", "_P")
                ElseIf InStr(strBrokerSymbol, "+") <> 0 Then
                    strReturn = Replace(strBrokerSymbol, "+", "_W")
                ElseIf InStr(strBrokerSymbol, "^") <> 0 Then
                    strReturn = Replace(strBrokerSymbol, "^", "_R")
                ElseIf InStr(strBrokerSymbol, "=") <> 0 Then
                    strReturn = Replace(strBrokerSymbol, "=", "_U")
                ElseIf IsAlpha(strBrokerSymbol, -1) Then
                    strReturn = strBrokerSymbol
                End If
                
                If (Len(strReturn) > 0) And (UCase(strBrokerExchange) = "BATS") Then
                    strReturn = strReturn & "@BATS"
                End If
            End If
            
        Case "CBOE"
            If InStr(strBrokerSymbol, ":") = 0 Then
                Set BrokerSym = m.Symbols.GetByBrokerSymbol("@")
                If Not BrokerSym Is Nothing Then
                    OptSym.FromRithmic strBrokerSymbol, "CBOE", BrokerSym.StrikeMult
                    strReturn = OptSym.ToGenesis
                End If
            End If
            
    End Select
    
    GenesisSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.GenesisSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerOrderType
'' Description: Attempt to convert the Genesis order type to the broker order type
'' Inputs:      Genesis Order Type
'' Returns:     Broker Order Type (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerOrderType(ByVal nGenesisOrderType As eTT_OrderType) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    Select Case nGenesisOrderType
        Case eTT_OrderType_Market
            strReturn = "Market"
        Case eTT_OrderType_Stop
            strReturn = "Stop"
        Case eTT_OrderType_Limit
            strReturn = "Limit"
        Case eTT_OrderType_StopWithLimit
            strReturn = "StopWithLimit"
        Case eTT_OrderType_MIT
            strReturn = "MarketIfTouched"
    End Select
    
    BrokerOrderType = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.BrokerOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddOrder
'' Description: Attempt to add the order
'' Inputs:      Order to add
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AddOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strOrderString As String        ' Broker order string
    
    bReturn = False
    strOrderString = BrokerOrderString(Order, True)
    
    If (Len(strOrderString) > 0) Then
        SendRithmicMessage eGDRithmicMessageType_AddOrder, strOrderString
        bReturn = True
    End If
    
    AddOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.AddOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AmendOrder
'' Description: Attempt to amend the order
'' Inputs:      Order to amend
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AmendOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strOrderString As String        ' Broker order string
    
    bReturn = False
    strOrderString = BrokerOrderString(Order, False)
    
    If (Len(strOrderString) > 0) Then
        SendRithmicMessage eGDRithmicMessageType_AmendOrder, strOrderString
        bReturn = True
    End If
    
    AmendOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.AmendOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrder
'' Description: Attempt to cancel the order
'' Inputs:      Order to cancel
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strOrderString As String        ' Broker order string
    
    bReturn = False
    strOrderString = BrokerOrderString(Order, False)
    
    If (Len(strOrderString) > 0) Then
        SendRithmicMessage eGDRithmicMessageType_CancelOrder, strOrderString
        bReturn = True
    End If
    
    CancelOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.CancelOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrder
'' Description: Attempt to park the order
'' Inputs:      Order to park
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ParkOrder(Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strOrderString As String        ' Broker order string
    
    bReturn = False
    strOrderString = BrokerOrderString(Order, False)
    
    If (Len(strOrderString) > 0) Then
        SendRithmicMessage eGDRithmicMessageType_CancelOrder, strOrderString
        bReturn = True
    End If
    
    ParkOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.ParkOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Refresh
'' Description: Ask the server for accounts, orders, fills, and positions
'' Inputs:      Verbose Refresh?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Refresh(Optional ByVal bVerbose As Boolean = False)
On Error GoTo ErrSection:

    GetAccounts

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.Refresh"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccounts
'' Description: Request accounts from the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAccounts()
On Error GoTo ErrSection:

    SendRithmicMessage eGDRithmicMessageType_GetAccounts, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.GetAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrders
'' Description: Request orders from the server
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetOrders(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim astrAccounts As cGdArray        ' Accounts array
    Dim lIndex As Long                  ' Index into a for loop

    If Len(strAccount) > 0 Then
        SendRithmicMessage eGDRithmicMessageType_GetOrders, strAccount
    Else
        Set astrAccounts = m.Broker.Accounts.MakeCopy
        For lIndex = 0 To astrAccounts.Size - 1
            m.OrdersRequestedFor.Add 0, astrAccounts(lIndex)
            SendRithmicMessage eGDRithmicMessageType_GetOrders, astrAccounts(lIndex)
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.GetOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetFills
'' Description: Request fills from the server
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetFills(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim astrAccounts As cGdArray        ' Accounts array
    Dim lIndex As Long                  ' Index into a for loop

    If Len(strAccount) > 0 Then
        SendRithmicMessage eGDRithmicMessageType_GetFills, strAccount
    Else
        Set astrAccounts = m.Broker.Accounts.MakeCopy
        For lIndex = 0 To astrAccounts.Size - 1
            m.FillsRequestedFor.Add 0, astrAccounts(lIndex)
            SendRithmicMessage eGDRithmicMessageType_GetFills, astrAccounts(lIndex)
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.GetFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositions
'' Description: Request positions from the server
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPositions(Optional ByVal strAccount As String = "")
On Error GoTo ErrSection:

    Dim astrAccounts As cGdArray        ' Accounts array
    Dim lIndex As Long                  ' Index into a for loop

    If Len(strAccount) > 0 Then
        SendRithmicMessage eGDRithmicMessageType_GetPositions, strAccount
    Else
        Set astrAccounts = m.Broker.Accounts.MakeCopy
        For lIndex = 0 To astrAccounts.Size - 1
            m.PositionsRequestedFor.Add 0, astrAccounts(lIndex)
            SendRithmicMessage eGDRithmicMessageType_GetPositions, astrAccounts(lIndex)
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetTradeRoutes
'' Description: Request trade routes from the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetTradeRoutes()
On Error GoTo ErrSection:

    SendRithmicMessage eGDRithmicMessageType_GetTradeRoutes, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.GetTradeRoutes"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetContractInfo
'' Description: Request contract information from the server
'' Inputs:      Genesis Symbol, Show Message?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetContractInfo(ByVal strGenesisSymbol As String, Optional ByVal bShowMessage As Boolean = False)
On Error GoTo ErrSection:

    Dim strBrokerSymbol As String       ' Broker symbol to request information for
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object

    strBrokerSymbol = BrokerSymbol(ConvertToTradeSymbol(strGenesisSymbol, CurrentTime), BrokerSym)
    
    If (Len(strBrokerSymbol) > 0) And (Not BrokerSym Is Nothing) Then
        SendRithmicMessage eGDRithmicMessageType_GetContractInfo, strBrokerSymbol & vbTab & BrokerSym.BrokerExchange
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.GetContractInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NextGenesisID
'' Description: Determine the next unique Genesis ID for the given account
'' Inputs:      Account Number
'' Returns:     Next Unique Genesis ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextGenesisID(ByVal strAccountNumber As String) As String
On Error GoTo ErrSection:

    Dim lDate As Long                   ' Date from the ini file
    Dim lCounter As Long                ' Counter from the ini file
    Dim lCustomerID As Long             ' Customer ID from the registry
    
    lDate = GetIniFileProperty("Date", 0&, "ID", m.Broker.IniFile)
    lCounter = GetIniFileProperty("Counter", 0&, "ID", m.Broker.IniFile)
    lCustomerID = (RI_GetLastDataServiceID \ 1000) Mod 1000000
    
    If lDate <> Date Then
        lDate = Date
        lCounter = 0&
    Else
        lCounter = lCounter + 1
    End If

    SetIniFileProperty "Date", lDate, "ID", m.Broker.IniFile
    SetIniFileProperty "Counter", lCounter, "ID", m.Broker.IniFile
    
    NextGenesisID = Format(lCustomerID, "000000") & Format(lDate, "00000") & Format(lCounter, "00000")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.NextGenesisID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StartingGenesisIdForOptNav
'' Description: Determine where Option Navigator should start its Genesis ID
'' Inputs:      None
'' Returns:     Starting ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StartingGenesisIdForOptNav() As String
On Error GoTo ErrSection:

    Dim lCustomerID As Long             ' Customer ID from the registry
    
    lCustomerID = (RI_GetLastDataServiceID \ 1000) Mod 1000000
    
    StartingGenesisIdForOptNav = Format(lCustomerID, "000000") & Format(Date, "00000") & "90000"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.StartingGenesisIdForOptNav"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsTradeableSymbol
'' Description: Is the given symbol in the conversion table?
'' Inputs:      Symbol
'' Returns:     True if tradeable, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsTradeableSymbol(ByVal strGenesisSymbol As String) As Boolean
On Error GoTo ErrSection:

    IsTradeableSymbol = (Len(BrokerSymbol(strGenesisSymbol)) > 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.IsTradeableSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsEnabledSymbol
'' Description: Is the given symbol enabled for trading for the user?
'' Inputs:      Symbol, Broker Base, Broker Exchange
'' Returns:     True if enabled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEnabledSymbol(ByVal strGenesisSymbol As String, Optional strBrokerBase As String, Optional strBrokerExchange As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    
    bReturn = (Len(BrokerSymbol(strGenesisSymbol, BrokerSym)) > 0)
    If bReturn Then
        strBrokerBase = BrokerSym.BrokerBase
        strBrokerExchange = BrokerSym.BrokerExchange
    End If
    
    IsEnabledSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.IsEnabledSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderTypeAllowed
'' Description: Is the given order type allowed for the given symbol?
'' Inputs:      Order Type, Symbol
'' Returns:     True if allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderTypeAllowed(ByVal nOrderType As eTT_OrderType, ByVal strGenesisSymbol As String) As Boolean
On Error GoTo ErrSection:

    Dim BrokerSym As cBrokerSymbol      ' Brokers symbol information
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Len(BrokerSymbol(strGenesisSymbol, BrokerSym)) > 0 Then
        If Not BrokerSym Is Nothing Then
            bReturn = BrokerSym.OrderTypeAllowed(nOrderType)
        End If
    End If
    
    OrderTypeAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.OrderTypeAllowed"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimeInForceAllowed
'' Description: Is the given time in force allowed for the given symbol?
'' Inputs:      Time In Force, Symbol
'' Returns:     True if allowed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TimeInForceAllowed(ByVal nTimeInForce As eTT_TimeInForce, ByVal strGenesisSymbol As String) As Boolean
On Error GoTo ErrSection:

    Dim BrokerSym As cBrokerSymbol      ' Brokers symbol information
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If Len(BrokerSymbol(strGenesisSymbol, BrokerSym)) > 0 Then
        If Not BrokerSym Is Nothing Then
            bReturn = BrokerSym.TifAllowed(nTimeInForce)
        End If
    End If
    
    TimeInForceAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.TimeInForceAllowed"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateToDouble
'' Description: Convert a string version of a date/time to a Julian double
'' Inputs:      String Date/Time
'' Returns:     Julian Date/Time
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DateToDouble(ByVal strDateTime As String) As Double
On Error GoTo ErrSection:

    Dim strDate As String               ' Date portion of the string
    Dim strTime As String               ' Time portion of the string
    Dim dReturn As Double               ' Return value of the function
    
    dReturn = 0#
    If Len(strDateTime) > 0 Then
        strDate = Parse(strDateTime, " ", 1)
        strTime = Parse(strDateTime, " ", 2)
        
        If Len(strTime) = 6 Then
            strTime = Left(strTime, 2) & ":" & Mid(strTime, 3, 2) & ":" & Right(strTime, 2)
        End If
        
        dReturn = JulFromLong(CLng(Val(strDate))) + CDate(strTime)
    End If
    
    DateToDouble = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.DateToDouble"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendRithmicMessage
'' Description: Send an App Mail message to the GenRithmic program
'' Inputs:      Type of Message, Message, Send Now?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SendRithmicMessage(ByVal nType As eGDRithmicMessageTypes, ByVal strMessage As String, Optional ByVal bSendNow As Boolean = True)
On Error GoTo ErrSection:

    Dim astrMessage As New cGdArray     ' Message split out into an array

    If nType = eGDRithmicMessageType_Connect Then
        m.Broker.DumpDebug vbTab & "Sending " & m.Broker.BrokerName & " Message (" & MessageTypeToString(nType) & "): " & ConnectStringLog
    Else
        m.Broker.DumpDebug vbTab & "Sending " & m.Broker.BrokerName & " Message (" & MessageTypeToString(nType) & "): " & strMessage
    End If
    
    frmOnlineBroker.gdBroker.CreateMessage m.Broker.ControlID, nType, strMessage, , bSendNow

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.SendRithmicMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle an incoming App Mail message from the GenRithmic program
'' Inputs:      Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal nType As eGDRithmicMessageTypes, ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim lAccountID As Long              ' Account the refresh is for
    Dim lIndex As Long                  ' Index into a for loop
    Dim strAccount As String            ' Account

    If Not m.Broker Is Nothing Then
        If (nType <> eGDRithmicMessageType_ConnectionInfo) And (nType <> eGDRithmicMessageType_Heartbeat) Then
            m.Broker.DumpDebug vbTab & m.Broker.BrokerName & " Message Received (" & MessageTypeToString(nType) & "): " & strMessage
        End If
        
        Select Case nType
            Case eGDRithmicMessageType_ConnectionInfo
                If Len(Parse(strMessage, vbTab, 2)) > 0 Then
                    m.Broker.HandleConnectionInfo CLng(Val(Parse(strMessage, vbTab, 2))), Parse(strMessage, vbTab, 3), m.strUserID
                End If
            
            Case eGDRithmicMessageType_AppLoaded
                m.Broker.AppLoaded = True
            
            Case eGDRithmicMessageType_AppUnloaded
                m.Broker.AppLoaded = False
            
            Case eGDRithmicMessageType_Heartbeat
                m.Broker.HeartBeat = gdTickCount
            
            Case eGDRithmicMessageType_AccountR
                m.Broker.AccountReceived strMessage
            
            Case eGDRithmicMessageType_Order
                m.Broker.OrderReceived strMessage, False
            
            Case eGDRithmicMessageType_OrderR
                If UCase(Parse(strMessage, vbTab, 1)) = "BEGIN" Then
                    If m.bOrdersBeginSent = False Then
                        m.bOrdersBeginSent = True
                        m.Broker.OrderReceived "BEGIN", True
                    End If
                ElseIf UCase(Parse(strMessage, vbTab, 1)) = "END" Then
                    m.OrdersRequestedFor.Remove Parse(strMessage, vbTab, 2)
                    If m.OrdersRequestedFor.Count = 0 Then
                        m.Broker.OrderReceived "END", True
                        m.bOrdersBeginSent = False
                    End If
                Else
                    m.Broker.OrderReceived strMessage, True
                End If
                
            Case eGDRithmicMessageType_Fill
                m.Broker.FillReceived strMessage, False
                
            Case eGDRithmicMessageType_FillR
                If UCase(Parse(strMessage, vbTab, 1)) = "BEGIN" Then
                    lAccountID = g.Broker.AccountIDForNumber(Parse(strMessage, vbTab, 2))
                    
                    For lIndex = m.CarPosFills.Count To 1 Step -1
                        If m.CarPosFills(lIndex).AccountID = lAccountID Then
                            m.CarPosFills.Remove lIndex
                        End If
                    Next lIndex
                    
                    If m.bFillsBeginSent = False Then
                        m.bFillsBeginSent = True
                        m.Broker.FillReceived "BEGIN", True
                    End If
                ElseIf UCase(Parse(strMessage, vbTab, 1)) = "END" Then
                    m.FillsRequestedFor.Remove Parse(strMessage, vbTab, 2)
                    If m.FillsRequestedFor.Count = 0 Then
                        m.Broker.FillReceived "END", True
                        m.bFillsBeginSent = False
                    End If
                ElseIf UCase(Parse(strMessage, vbTab, 1)) = "SOD" Then
                    HandleCarriedPositionFill strMessage
                Else
                    m.Broker.FillReceived strMessage, True
                End If
                
            Case eGDRithmicMessageType_Position
                m.Broker.PositionReceived strMessage, False
            
            Case eGDRithmicMessageType_PositionR
                If UCase(Parse(strMessage, vbTab, 1)) = "BEGIN" Then
                    If m.bPositionsBeginSent = False Then
                        m.bPositionsBeginSent = True
                        m.Broker.PositionReceived "BEGIN", True
                    End If
                ElseIf UCase(Parse(strMessage, vbTab, 1)) = "END" Then
                    strAccount = Parse(strMessage, vbTab, 2)
                    SendFlatPositions strAccount
                    
                    m.PositionsRequestedFor.Remove strAccount
                    If m.PositionsRequestedFor.Count = 0 Then
                        m.Broker.PositionReceived "END", True
                        m.bPositionsBeginSent = False
                    End If
                Else
                    m.Broker.PositionReceived strMessage, True
                End If
                
            Case eGDRithmicMessageType_ExchList
                HandleExchangeList strMessage
                
            Case eGDRithmicMessageType_TradeRoute
                HandleTradeRoute strMessage
                            
        End Select
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.HandleMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountFromInfo
'' Description: Fill an account object from the given message
'' Inputs:      Message, Account
'' Returns:     True if OK, False otherwise
''
'' Fields:      Account, FCM, IB
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountFromInfo(ByVal strMessage As String, Acct As cPtAccount) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrAccount As cGdArray         ' Account information broken up into an array
    Dim strIbInfo As String             ' Introducing broker info
    Dim lPos As Long                    ' Position in the array

    bReturn = False
    If Len(strMessage) > 0 Then
        Set astrAccount = New cGdArray
        astrAccount.SplitFields strMessage, vbTab

        If UCase(astrAccount(0)) = "BEGIN" Then
            m.astrIbs.Clear
            m.IbIdsByAccount.Clear
        ElseIf UCase(astrAccount(0)) = "END" Then
            GetTradeRoutes
        Else
            With Acct
                .AccountNumber = astrAccount(0)
                .AccountType = m.Broker.Broker
                .Broker = ""
                .ClosedProfit = 0#
                .Comms = 0#
                .CurrentBalance = 0#
                .FillRT = False
                .HoldOcoAtBroker = False
                .Name = .AccountNumber
                .Password = m.strPassword
                .SecTypeMask = 31
                .StartingBalance = 0#
                .StartingDate = Date
                .UserName = m.strUserID
            End With
            
            strIbInfo = astrAccount(1) & vbTab & astrAccount(2)
            If m.astrIbs.BinarySearch(strIbInfo, lPos) = False Then
                m.astrIbs.Add strIbInfo, lPos
            End If
            If m.IbIdsByAccount.Exists(astrAccount(0)) Then
                m.IbIdsByAccount(astrAccount(0)) = strIbInfo
            Else
                m.IbIdsByAccount.Add strIbInfo, astrAccount(0)
            End If
            
            bReturn = True
        End If
    End If
    
    AccountFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.AccountFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderFromInfo
'' Description: Fill an order object from the given message
'' Inputs:      Message, Order
'' Returns:     True if OK, False otherwise
''
'' Fields:      Genesis ID, Order ID, Account, Symbol, Exchange, Order Type
''              Side, Quantity, Limit, Stop, Time In Force, Remaining Quantity,
''              Cumulative Quantity, Order Time, Status, Reason
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderFromInfo(ByVal strMessage As String, Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrOrder As cGdArray           ' Order information split out from the message
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim lQuantityFilled As Long         ' Quantity of the order that has filled
    Dim Bars As New cGdBars             ' Bars object
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set astrOrder = New cGdArray
            astrOrder.SplitFields strMessage, vbTab
            
            strGenesisSymbol = GenesisSymbol(astrOrder(3), astrOrder(4), BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
                SetBarProperties Bars, strGenesisSymbol
                
                Set Order = New cPtOrder
                With Order
                    .AccountID = g.Broker.AccountIDForNumber(astrOrder(2))
                    .BrokerID = astrOrder(1)
                    .Buy = (Left(UCase(astrOrder(6)), 3) = "BUY")
                    .GenesisOrderID = astrOrder(0)
                    .LimitPrice = GenesisPrice(astrOrder(8), BrokerSym)
                    .Message = astrOrder(15)
                    .OrderType = GenesisOrderType(astrOrder(5))
                    .Quantity = CLng(Val(astrOrder(7)))
                    .Status = GenesisOrderStatus(astrOrder(14))
                    If Len(astrOrder(13)) = 0 Then
                        .StatusDate = ConvertToBrokerDate(CurrentTime, m.Broker.Broker, strGenesisSymbol)
                    Else
                        .StatusDate = DateToDouble(astrOrder(13))
                    End If
                    .SessionDate = Bars.SessionDateForTradeTime(ConvertBrokerDate(.StatusDate, m.Broker.Broker, strGenesisSymbol, False))
                    .StopPrice = GenesisPrice(astrOrder(9), BrokerSym)
                    .SymbolOrSymbolID = strGenesisSymbol
                    lQuantityFilled = CLng(Val(astrOrder(12)))
                    
                    If .OrderType = eTT_OrderType_MIT Then
                        .LimitPrice = .StopPrice
                        .StopPrice = 0#
                    End If
                    
                    Select Case UCase(astrOrder(10))
                        Case "DAY"
                            .Expiration = .SessionDate * -1&
                        Case "GTC"
                            .Expiration = 0
                    End Select
                    
                    If lQuantityFilled > 0 Then
                        If (.Status = eTT_OrderStatus_Working) Then
                            .Status = eTT_OrderStatus_Partial
                        ElseIf (.Status = eTT_OrderStatus_Cancelled) Then
                            .Status = eTT_OrderStatus_BalCancelled
                        End If
                    End If
                End With
                
                bReturn = True
            End If
        End If
    End If
    
    OrderFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.OrderFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillFromInfo
'' Description: Fill a fill object from the given message
'' Inputs:      Message, Fill, Genesis Order ID, Carried?
'' Returns:     True if OK, False otherwise
''
'' Fields:      Genesis ID, Order ID, Fill ID, Account, Symbol, Exchange,
''              B/S, Fill Quantity, Fill Price, Fill Time
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillFromInfo(ByVal strMessage As String, Fill As cPtFill, Optional strGenesisOrderID As String = "", Optional ByVal bCarried As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrFill As cGdArray            ' Fill information split out from the message
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim Bars As New cGdBars             ' Bars object
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set astrFill = New cGdArray
            astrFill.SplitFields strMessage, vbTab
            
            strGenesisSymbol = GenesisSymbol(astrFill(4), astrFill(5), BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
                SetBarProperties Bars, strGenesisSymbol
                
                Set Fill = New cPtFill
                With Fill
                    .AccountID = g.Broker.AccountIDForNumber(astrFill(3))
                    .BrokerID = astrFill(2)
                    .BrokerOrderID = astrFill(1)
                    .Buy = (Left(UCase(astrFill(6)), 3) = "BUY")
                    .Fees = 0
                    .FillDate = DateToDouble(astrFill(9))
                    .Price = GenesisPrice(astrFill(8), BrokerSym)
                    .Quantity = CLng(Val(astrFill(7)))
                    .SessionDate = Bars.SessionDateForTradeTime(.FillDateConverted(""))
                    .SymbolOrSymbolID = strGenesisSymbol
                End With
                
                strGenesisOrderID = ""
                
                bReturn = True
            End If
        End If
    End If
    
    FillFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.FillFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CarriedFillFromInfo
'' Description: Fill a fill object from the given message
'' Inputs:      Message, Fill, Genesis Order ID
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CarriedFillFromInfo(ByVal strMessage As String, Fill As cPtFill, Optional strGenesisOrderID As String = "") As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            bReturn = FillFromInfo(strMessage, Fill, strGenesisOrderID, True)
        End If
    End If
    
    CarriedFillFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.CarriedFillFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionFromInfo
'' Description: Fill a position object from the given message
'' Inputs:      Message, Position
'' Returns:     True if OK, False otherwise
''
'' Fields:      Account, Symbol, Exchange, Position, Realized, Unrealized,
''              Avg Entry, Carried Position, Carried Realized, Carried Unrealized,
''              Carried Avg Entry, Instrument Type, ID, Carried Info
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionFromInfo(ByVal strMessage As String, Position As cPtPosition) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrPosition As cGdArray        ' Position information split out from the message
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim astrFills As cGdArray           ' Array of fill information
    Dim lIndex As Long                  ' Index into a for loop
    Dim Fill As cPtFill                 ' Fill object
    Dim lPositionDate As Long           ' Date of the position
    Dim strKey As String                ' Key into the carried position collection
    Dim CarPosFill As cPtFill           ' Carried Position information
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set astrPosition = New cGdArray
            astrPosition.SplitFields strMessage, vbTab
            
            strGenesisSymbol = GenesisSymbol(astrPosition(1), astrPosition(2), BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
                Set Position = New cPtPosition
                With Position
                    .AccountID = g.Broker.AccountIDForNumber(astrPosition(0))
                    .SymbolOrSymbolID = strGenesisSymbol
                    .CurrentPosition = CLng(Val(astrPosition(3)))
                    
                    strKey = Str(.AccountID) & vbTab & .Symbol
                    If m.CarPosFills.Exists(strKey) Then
                        Set CarPosFill = m.CarPosFills(strKey)
                        If CarPosFill.Buy Then
                            .CarriedPosition = CarPosFill.Quantity
                        Else
                            .CarriedPosition = CarPosFill.Quantity * -1&
                        End If
                        
                        If CarPosFill.Quantity > 0 Then
                            .CarriedFills.Add CarPosFill
                        End If
                        
                        m.CarPosFills.Remove strKey
                    End If
                End With
                
                bReturn = True
            End If
        End If
    End If
    
    PositionFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.PositionFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbolList
'' Description: Return a list of the Genesis symbols that are in the file
'' Inputs:      Only if stream?
'' Returns:     List of symbols
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbolList(Optional ByVal bOnlyIfStream As Boolean = True) As cGdArray
On Error GoTo ErrSection:

    Set GenesisSymbolList = m.Symbols.GenesisSymbolList

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.GenesisSymbolList"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimeZone
'' Description: Determine the time zone this broker sends time in for the symbol
'' Inputs:      Symbol
'' Returns:     Time Zone
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TimeZone(ByVal strSymbol As String) As String
On Error GoTo ErrSection:

    TimeZone = "GMT"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.TimeZone"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    JustConnected
'' Description: Handle the fact that we have just gone connected with the broker
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub JustConnected()
On Error GoTo ErrSection:

    Dim astrSymbols As cGdArray         ' Array of symbols on forms
    Dim lIndex As Long                  ' Index into a for loop

    m.Broker.DumpDebug "Getting Contracts for Symbols on Forms"
    Set astrSymbols = g.Broker.SymbolsOnFormsForBroker(m.Broker.Broker)
    For lIndex = 0 To astrSymbols.Size - 1
        GetContractInfo astrSymbols(lIndex)
    Next lIndex

    GetAccounts

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.JustConnected"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolInformation
'' Description: Get the symbol information for the given symbol
'' Inputs:      Genesis Symbol
'' Returns:     Symbol Information
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SymbolInformation(ByVal strGenesisSymbol As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol information
    
    strReturn = ""
    If Len(BrokerSymbol(strGenesisSymbol, BrokerSym)) > 0 Then
        strReturn = BrokerSym.OrderTypeMask & vbTab & BrokerSym.TifMask
    End If
    
    SymbolInformation = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.SymbolInformation"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerPrice
'' Description: Convert a Genesis price to the broker price
'' Inputs:      Genesis Price, Broker Symbol Object
'' Returns:     Broker Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerPrice(ByVal dGenesisPrice As Double, ByVal BrokerSym As cBrokerSymbol) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = Str(dGenesisPrice)
    If Not BrokerSym Is Nothing Then
        strReturn = m.Broker.BrokerPrice(dGenesisPrice, BrokerSym.PriceMult, BrokerSym.PriceFormat)
    End If
    
    BrokerPrice = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.BrokerPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisPrice
'' Description: Attempt to convert the broker price to a Genesis price
'' Inputs:      Broker Price, Broker Symbol Object
'' Returns:     Genesis Price
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisPrice(ByVal strBrokerPrice As String, ByVal BrokerSym As cBrokerSymbol) As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the function
    
    dReturn = 0#
    If Not BrokerSym Is Nothing Then
        dReturn = m.Broker.GenesisPrice(strBrokerPrice, BrokerSym.PriceMult)
    End If
    
    GenesisPrice = dReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.GenesisPrice"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderType
'' Description: Convert a broker order type to a Genesis order type
'' Inputs:      Broker order type
'' Returns:     Genesis Order Type (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderType(ByVal strBrokerOrderType As String) As eTT_OrderType
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderType        ' Return value for the function
    
    Select Case UCase(strBrokerOrderType)
        Case "LIMIT"
            nReturn = eTT_OrderType_Limit
        Case "MARKET"
            nReturn = eTT_OrderType_Market
        Case "STOP"
            nReturn = eTT_OrderType_Stop
        Case "STOPWITHLIMIT"
            nReturn = eTT_OrderType_StopWithLimit
        Case "MARKETIFTOUCHED"
            nReturn = eTT_OrderType_MIT
        Case "LIMITIFTOUCHED"
    End Select
    
    GenesisOrderType = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.GenesisOrderType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisOrderStatus
'' Description: Convert a broker order status to a Genesis order status
'' Inputs:      Broker order status
'' Returns:     Genesis order status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisOrderStatus(ByVal strBrokerOrderStatus As String) As eTT_OrderStatus
On Error GoTo ErrSection:

    Dim nReturn As eTT_OrderStatus      ' Return value for the function

    Select Case UCase(strBrokerOrderStatus)
        Case "CANCELREJECTED":          ' Cancel has been rejected
            nReturn = eTT_OrderStatus_Rejected
        Case "CANCELPENDING":           ' Cancel is pending
            nReturn = eTT_OrderStatus_CancelPending
        Case "CANCELRECDGWAY":          ' Exchange gateway has received cancel request
            nReturn = eTT_OrderStatus_CancelPending
        Case "CANCELRECEIVED":          ' Cancel has been received by the Rithmic servers
            nReturn = eTT_OrderStatus_CancelPending
        Case "CANCELSENTEXCH":          ' Cancel has been sent to the exchange
            nReturn = eTT_OrderStatus_CancelPending
        Case "CANCELLED":               ' Order has been cancelled
            nReturn = eTT_OrderStatus_Cancelled
        Case "FAILURE":                 ' Order failed
            nReturn = eTT_OrderStatus_Error
        Case "FILLED":                  ' Order has been filled
            nReturn = eTT_OrderStatus_Filled
        Case "BALANCECANCELLED":        ' Balance of the order has been cancelled
            nReturn = eTT_OrderStatus_BalCancelled
        Case "REJECTED":                 ' Order has been rejected"
            nReturn = eTT_OrderStatus_Rejected
        Case "COMPLETEUNKNOWN":         ' Order is complete for an unknown reason
            nReturn = eTT_OrderStatus_Error
        Case "AMENDTRIGGERPENDING":     ' Order is pending an amendment trigger
            nReturn = eTT_OrderStatus_Working
        Case "AMENDED":                 ' Order has been amended
            nReturn = eTT_OrderStatus_Working
        Case "AMENDREJECTED":           ' Amendment has been rejected
            nReturn = eTT_OrderStatus_Rejected
        Case "AMENDPENDING":             ' Amendment is pending"
            nReturn = eTT_OrderStatus_AmendPending
        Case "AMENDRECDGWAY":           ' Exchange gateway has received order amendment
            nReturn = eTT_OrderStatus_AmendPending
        Case "AMENDRECEIVED":           ' Broker has recieved the order amendment
            nReturn = eTT_OrderStatus_AmendPending
        Case "AMENDSENTEXCH":           ' Amendment has been sent to the exchange
            nReturn = eTT_OrderStatus_AmendPending
        Case "WORKING":                 ' Order is working
            nReturn = eTT_OrderStatus_Working
        Case "SENT":                    ' Order is pending
            nReturn = eTT_OrderStatus_Sent
        Case "RECDGWAY":                ' Order has been received by the exchange gateway
            nReturn = eTT_OrderStatus_Sent
        Case "RECEIVED":                ' Broker has received the order
            nReturn = eTT_OrderStatus_Sent
        Case "SENTEXCH":                ' Order has been sent to the exchange
            nReturn = eTT_OrderStatus_Sent
        Case "RELEASED":                ' Order has been released by the ORM
        Case "RELEASEPENDING":          ' Order's release by the ORM is pending
        Case "PARTIALFILLED":           ' Order has partially filled
            nReturn = eTT_OrderStatus_Partial
        Case "SENTORM":                 ' Order has been sent to the ORM
        Case "TRIGGERPENDING":          ' Order is pending a trigger
            nReturn = eTT_OrderStatus_Working
        Case "ERROR":                   ' Error sending order
            nReturn = eTT_OrderStatus_Error
        
    End Select
    
    GenesisOrderStatus = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.GenesisOrderStatus"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerBuySell
'' Description: Convert a Genesis Buy/Sell to the broker buy/sell
'' Inputs:      Genesis Buy/Sell
'' Returns:     Broker Buy/Sell
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerBuySell(ByVal bBuySell As Boolean, ByVal Order As cPtOrder) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strSecType As String            ' Security type for the symbol passed in
    Dim lCurPos As Long                 ' Current position
    
    If bBuySell Then
        strReturn = "Buy"
    Else
        strSecType = SecurityType(Order.Symbol, True)
        If strSecType = "S" Then
            lCurPos = m.Broker.CurrentPosition(g.Broker.AccountNumberForID(Order.AccountID), Order.Symbol)
            If lCurPos > 0 Then
                If Order.Quantity <= lCurPos Then
                    strReturn = "Sell"
                Else
                    strReturn = "SellShortExempt"
                End If
            Else
                strReturn = "SellShortExempt"
            End If
        Else
            strReturn = "Sell"
        End If
    End If
    
    BrokerBuySell = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.BrokerBuySell"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerTIF
'' Description: Convert a Genesis Time-in-Force to the broker Time-in-Force
'' Inputs:      Genesis Time-in-Force
'' Returns:     Broker Time-in-Force
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerTIF(ByVal nGenesisTIF As eTT_TimeInForce) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    Select Case nGenesisTIF
        Case eTT_TimeInForce_Day
            strReturn = "Day"
        Case eTT_TimeInForce_GTC
            strReturn = "GTC"
        Case eTT_TimeInForce_GTD
            strReturn = "GTD"
    End Select
    
    BrokerTIF = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.BrokerTIF"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerOrderString
'' Description: Build an order string to send to the broker
'' Inputs:      Order, New Order?
'' Returns:     Broker Order String
''
'' Fields:      GenesisID, BrokerID, Account, Symbol, Symbol, Order Type,
''              Buy/Sell, Quantity, Limit, Stop, Reserved, Reserved, Exchange,
''              Reserved, TIF, Reserved, Reserved, Reserved, Reserved,
''              Entry Type, Trade Route
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerOrderString(ByVal Order As cPtOrder, ByVal bNewOrder As Boolean) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrOrder As cGdArray           ' Array of order information to send to broker
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim strBrokerSymbol As String       ' Broker symbol
    
    strReturn = ""
    strBrokerSymbol = BrokerSymbol(Order.Symbol, BrokerSym)
    
    If (Len(strBrokerSymbol) > 0) And (Not BrokerSym Is Nothing) Then
        Set astrOrder = New cGdArray
        astrOrder.Create eGDARRAY_Strings, 17
        
        astrOrder(0) = Order.GenesisOrderID
        If bNewOrder Then
            astrOrder(1) = ""
        Else
            astrOrder(1) = Order.BrokerID
        End If
        astrOrder(2) = g.Broker.AccountNumberForID(Order.AccountID)
        astrOrder(3) = strBrokerSymbol
        astrOrder(4) = strBrokerSymbol
        astrOrder(5) = BrokerOrderType(Order.OrderType)
        astrOrder(6) = BrokerBuySell(Order.Buy, Order)
        astrOrder(7) = Str(Order.Quantity)
        astrOrder(8) = BrokerPrice(Order.LimitPrice, BrokerSym)
        astrOrder(9) = BrokerPrice(Order.StopPrice, BrokerSym)
        astrOrder(10) = ""
        astrOrder(11) = ""
        astrOrder(12) = BrokerSym.BrokerExchange
        astrOrder(13) = ""
        astrOrder(14) = BrokerTIF(Order.TimeInForce)
        astrOrder(15) = ""
        astrOrder(16) = ""
        astrOrder(17) = ""
        astrOrder(18) = ""
        If Order.IsAutomated Then
            astrOrder(19) = "A"
        Else
            astrOrder(19) = "M"
        End If
        astrOrder(20) = FindTradeRoute(BrokerSym.BrokerExchange, astrOrder(2))
        
        strReturn = astrOrder.JoinFields(vbTab)
    End If
    
    BrokerOrderString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.BrokerOrderString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleCarriedPositionFill
'' Description: Handle a carried position fill
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleCarriedPositionFill(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim Fill As cPtFill                 ' Fill object to represent the carried position
    Dim strKey As String                ' Key into the collection
    
    If FillFromInfo(strMessage, Fill) Then
        strKey = Str(Fill.AccountID) & vbTab & Fill.Symbol
        If m.CarPosFills.Exists(strKey) Then
            Set m.CarPosFills(strKey) = Fill
        Else
            m.CarPosFills.Add Fill, strKey
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.HandleCarriedPositionFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SendFlatPositions
'' Description: Send positions that are flat, but have carried fills
'' Inputs:      Account
'' Returns:     None
''
'' Fields:      Account, Symbol, Exchange, Position, Realized, Unrealized,
''              Avg Entry, Carried Position, Carried Realized, Carried Unrealized,
''              Carried Avg Entry, Instrument Type, ID, Carried Info
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SendFlatPositions(ByVal strAccount As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lAccountID As Long              ' Account ID for the account passed in
    Dim Fill As cPtFill                 ' Fill stored in the collection
    Dim strBrokerSymbol As String       ' Broker symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim astrFields As cGdArray          ' Array of fields
    
    lAccountID = g.Broker.AccountIDForNumber(strAccount)
    Set astrFields = New cGdArray
    astrFields.Create eGDARRAY_Strings
    
    ' DAJ 02/01/2012: Need to count backwards since the m.Broker.PositionReceived is going
    ' to end up removing the carried position from the collection...
    For lIndex = m.CarPosFills.Count To 1 Step -1
        Set Fill = m.CarPosFills(lIndex)
        If Fill.AccountID = lAccountID Then
            strBrokerSymbol = BrokerSymbol(Fill.Symbol, BrokerSym)
            If Len(strBrokerSymbol) > 0 Then
                astrFields(0) = strAccount
                astrFields(1) = strBrokerSymbol
                astrFields(2) = BrokerSym.BrokerExchange
                astrFields(3) = "0"
                
                m.Broker.DumpDebug vbTab & "Flat Position Created: " & astrFields.JoinFields(vbTab)
                m.Broker.PositionReceived astrFields.JoinFields(vbTab), True
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.SendFlatPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLoginInfo
'' Description: Get login information from the user
'' Inputs:      User Name, Switching
'' Returns:     True if user OKed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetLoginInfo(Optional ByVal strUserName As String = "", Optional ByVal bSwitching As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lSystemID As Long               ' System to connect to
    Dim strCnctFile As String           ' Connection information file

    bReturn = frmRithmicLogin.ShowMe(m.Broker, strUserName, bSwitching)
    If bReturn = True Then
        m.strUserID = frmRithmicLogin.UserID
        m.strPassword = frmRithmicLogin.Password
        
        lSystemID = frmRithmicLogin.SystemID
        strCnctFile = GetIniFileProperty("CnctFile", "", "Environment" & Str(lSystemID), m.Broker.ConnectIni)
        If Len(strCnctFile) > 0 Then
            m.strCnctFile = AddSlash(App.Path) & "Provided\" & strCnctFile
        Else
            m.strCnctFile = m.Broker.ConnectIni
        End If
        
        m.Broker.AskedPassword = True
    End If
    
    GetLoginInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.GetLoginInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleExchangeList
'' Description: Handle the exchange list coming from Rithmic
'' Inputs:      Exchange List
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleExchangeList(ByVal strExchangeList As String)
On Error GoTo ErrSection:

    m.astrExchanges.Clear
    m.astrExchanges.SplitFields strExchangeList, ","
    m.astrExchanges.Sort

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.HandleExchangeList"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleTradeRoute
'' Description: Handle a trade route coming from Rithmic
'' Inputs:      Trade Route
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleTradeRoute(ByVal strTradeRoute As String)
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Array of fields from the given string
    Dim TradeRoute As cTradeRoute       ' Trade route object
    
    If Len(strTradeRoute) > 0 Then
        Set astrFields = New cGdArray
        astrFields.SplitFields strTradeRoute, vbTab
        
        If UCase(astrFields(0)) = "BEGIN" Then
            m.TradeRoutes.Clear
        ElseIf UCase(astrFields(0)) = "END" Then
            SubscribeTradeRoutes
        Else
            Set TradeRoute = New cTradeRoute
            TradeRoute.FromRithmicString strTradeRoute
            
            If m.TradeRoutes.Exists(TradeRoute.Key) Then
                m.TradeRoutes(TradeRoute.Key) = m.TradeRoutes(TradeRoute.Key) & "," & TradeRoute.TradeRoute
            Else
                m.TradeRoutes.Add TradeRoute.TradeRoute, TradeRoute.Key
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.HandleTradeRoute"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubscribeTradeRoutes
'' Description: Subscribe to the trade routes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SubscribeTradeRoutes()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 0 To m.astrIbs.Size - 1
        SendRithmicMessage eGDRithmicMessageType_SubscribeTradeRoute, m.astrIbs(lIndex)
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.SubscribeTradeRoute"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FindTradeRoute
'' Description: Determine the trade route name for the given information
'' Inputs:      Exchange, Account
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindTradeRoute(ByVal strExchange As String, ByVal strAccountNumber As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strIbId As String               ' FCM/IB information for the account
    Dim strKey As String                ' Key into the collection
    Dim astrTradeRoutes As cGdArray     ' Trade routes for the key
    
    strReturn = "electronic"

    If m.IbIdsByAccount.Exists(strAccountNumber) Then
        strIbId = m.IbIdsByAccount(strAccountNumber)
        
        strKey = strExchange & vbTab & strIbId
        If m.TradeRoutes.Exists(strKey) Then
            Set astrTradeRoutes = New cGdArray
            astrTradeRoutes.SplitFields m.TradeRoutes(strKey), ","
            
            If astrTradeRoutes.Size = 2 Then
                If UCase(astrTradeRoutes(0)) = "ELECTRONIC" Then
                    strReturn = astrTradeRoutes(1)
                Else
                    strReturn = astrTradeRoutes(0)
                End If
            ElseIf astrTradeRoutes.Size > 0 Then
                strReturn = astrTradeRoutes(0)
            End If
        End If
    End If
    
    FindTradeRoute = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.FindTradeRoute"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MessageTypeToString
'' Description: Convert a message type enumeration to a string
'' Inputs:      Message Type
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MessageTypeToString(ByVal nType As eGDRithmicMessageTypes) As String
On Error GoTo ErrSection:

    Select Case nType
        Case eGDRithmicMessageType_Connect
            MessageTypeToString = "eGDRithmicMessageType_Connect"
        Case eGDRithmicMessageType_Disconnect
            MessageTypeToString = "eGDRithmicMessageType_Disconnect"
        Case eGDRithmicMessageType_AddOrder
            MessageTypeToString = "eGDRithmicMessageType_AddOrder"
        Case eGDRithmicMessageType_AmendOrder
            MessageTypeToString = "eGDRithmicMessageType_AmendOrder"
        Case eGDRithmicMessageType_CancelOrder
            MessageTypeToString = "eGDRithmicMessageType_CancelOrder"
        Case eGDRithmicMessageType_UnloadApp
            MessageTypeToString = "eGDRithmicMessageType_UnloadApp"
        Case eGDRithmicMessageType_GetAccounts
            MessageTypeToString = "eGDRithmicMessageType_GetAccounts"
        Case eGDRithmicMessageType_GetOrders
            MessageTypeToString = "eGDRithmicMessageType_GetOrders"
        Case eGDRithmicMessageType_GetFills
            MessageTypeToString = "eGDRithmicMessageType_GetFills"
        Case eGDRithmicMessageType_GetPositions
            MessageTypeToString = "eGDRithmicMessageType_GetPositions"
        
        Case eGDRithmicMessageType_ConnectionInfo
            MessageTypeToString = "eGDRithmicMessageType_ConnectionInfo"
        Case eGDRithmicMessageType_AppLoaded
            MessageTypeToString = "eGDRithmicMessageType_AppLoaded"
        Case eGDRithmicMessageType_AppUnloaded
            MessageTypeToString = "eGDRithmicMessageType_AppUnloaded"
        Case eGDRithmicMessageType_Heartbeat
            MessageTypeToString = "eGDRithmicMessageType_Heartbeat"
        Case eGDRithmicMessageType_Order
            MessageTypeToString = "eGDRithmicMessageType_Order"
        Case eGDRithmicMessageType_AccountR
            MessageTypeToString = "eGDRithmicMessageType_AccountR"
        Case eGDRithmicMessageType_OrderR
            MessageTypeToString = "eGDRithmicMessageType_OrderR"
        Case eGDRithmicMessageType_FillR
            MessageTypeToString = "eGDRithmicMessageType_FillR"
        Case eGDRithmicMessageType_PositionR
            MessageTypeToString = "eGDRithmicMessageType_PositionR"
            
        Case eGDRithmicMessageType_Fill
            MessageTypeToString = "eGDRithmicMessageType_Fill"
        Case eGDRithmicMessageType_Position
            MessageTypeToString = "eGDRithmicMessageType_Position"
        Case eGDRithmicMessageType_ExchList
            MessageTypeToString = "eGDRithmicMessageType_ExchList"
        Case eGDRithmicMessageType_GetTradeRoutes
            MessageTypeToString = "eGDRithmicMessageType_GetTradeRoutes"
        Case eGDRithmicMessageType_TradeRoute
            MessageTypeToString = "eGDRithmicMessageType_TradeRoute"
        Case eGDRithmicMessageType_SubscribeTradeRoute
            MessageTypeToString = "eGDRithmicMessageType_SubscribeTradeRoute"
        Case eGDRithmicMessageType_UnsubscribeTradeRoute
            MessageTypeToString = "eGDRithmicMessageType_UnsubscribeTradeRoute"
        Case eGDRithmicMessageType_GetContractInfo
            MessageTypeToString = "eGDRithmicMessageType_GetContractInfo"
        Case eGDRithmicMessageType_ContractInfo
            MessageTypeToString = "eGDRithmicMessageType_ContractInfo"
            
        Case Else
            MessageTypeToString = Str(nType)
    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.MessageTypeToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StringToMessageType
'' Description: Convert a string to a message type enumeration
'' Inputs:      String
'' Returns:     Message Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StringToMessageType(ByVal strMessageType As String) As eGDRithmicMessageTypes
On Error GoTo ErrSection:
    
    Dim nReturn As eGDRithmicMessageTypes   ' Return value for the function
    
    Select Case strMessageType
        Case "eGDRithmicMessageType_Connect"
            nReturn = eGDRithmicMessageType_Connect
        Case "eGDRithmicMessageType_Disconnect"
            nReturn = eGDRithmicMessageType_Disconnect
        Case "eGDRithmicMessageType_AddOrder"
            nReturn = eGDRithmicMessageType_AddOrder
        Case "eGDRithmicMessageType_AmendOrder"
            nReturn = eGDRithmicMessageType_AmendOrder
        Case "eGDRithmicMessageType_CancelOrder"
            nReturn = eGDRithmicMessageType_CancelOrder
        Case "eGDRithmicMessageType_UnloadApp"
            nReturn = eGDRithmicMessageType_UnloadApp
        Case "eGDRithmicMessageType_GetAccounts"
            nReturn = eGDRithmicMessageType_GetAccounts
        Case "eGDRithmicMessageType_GetOrders"
            nReturn = eGDRithmicMessageType_GetOrders
        Case "eGDRithmicMessageType_GetFills"
            nReturn = eGDRithmicMessageType_GetFills
        Case "eGDRithmicMessageType_GetPositions"
            nReturn = eGDRithmicMessageType_GetPositions
        
        Case "eGDRithmicMessageType_ConnectionInfo"
            nReturn = eGDRithmicMessageType_ConnectionInfo
        Case "eGDRithmicMessageType_AppLoaded"
            nReturn = eGDRithmicMessageType_AppLoaded
        Case "eGDRithmicMessageType_AppUnloaded"
            nReturn = eGDRithmicMessageType_AppUnloaded
        Case "eGDRithmicMessageType_Heartbeat"
            nReturn = eGDRithmicMessageType_Heartbeat
        Case "eGDRithmicMessageType_Order"
            nReturn = eGDRithmicMessageType_Order
        Case "eGDRithmicMessageType_AccountR"
            nReturn = eGDRithmicMessageType_AccountR
        Case "eGDRithmicMessageType_OrderR"
            nReturn = eGDRithmicMessageType_OrderR
        Case "eGDRithmicMessageType_FillR"
            nReturn = eGDRithmicMessageType_FillR
        Case "eGDRithmicMessageType_PositionR"
            nReturn = eGDRithmicMessageType_PositionR
        
        Case "eGDRithmicMessageType_Fill"
            nReturn = eGDRithmicMessageType_Fill
        Case "eGDRithmicMessageType_Position"
            nReturn = eGDRithmicMessageType_Position
        Case "eGDRithmicMessageType_ExchList"
            nReturn = eGDRithmicMessageType_ExchList
        Case "eGDRithmicMessageType_GetTradeRoutes"
            nReturn = eGDRithmicMessageType_GetTradeRoutes
        Case "eGDRithmicMessageType_TradeRoute"
            nReturn = eGDRithmicMessageType_TradeRoute
        Case "eGDRithmicMessageType_SubscribeTradeRoute"
            nReturn = eGDRithmicMessageType_SubscribeTradeRoute
        Case "eGDRithmicMessageType_UnsubscribeTradeRoute"
            nReturn = eGDRithmicMessageType_UnsubscribeTradeRoute
        Case "eGDRithmicMessageType_GetContractInfo"
            nReturn = eGDRithmicMessageType_GetContractInfo
        Case "eGDRithmicMessageType_ContractInfo"
            nReturn = eGDRithmicMessageType_ContractInfo
        
        Case Else
            nReturn = 0
    End Select
    
    StringToMessageType = nReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRithmic.StringToMessageType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Do any initialization that needs to be done for the class
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Broker = New cBroker
    
    Set m.OrdersRequestedFor = New cGdTree
    m.bOrdersBeginSent = False
    Set m.FillsRequestedFor = New cGdTree
    m.bFillsBeginSent = False
    Set m.PositionsRequestedFor = New cGdTree
    m.bPositionsBeginSent = False
    
    Set m.astrExchanges = New cGdArray
    m.astrExchanges.Create eGDARRAY_Strings
    Set m.astrIbs = New cGdArray
    m.astrIbs.Create eGDARRAY_Strings
    Set m.IbIdsByAccount = New cGdTree
    Set m.TradeRoutes = New cGdTree
    
    Set m.CarPosFills = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Do any clean up that needs to be done for the class
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Broker = Nothing
    Set m.OrdersRequestedFor = Nothing
    Set m.FillsRequestedFor = Nothing
    Set m.PositionsRequestedFor = Nothing
    Set m.astrExchanges = Nothing
    Set m.astrIbs = Nothing
    Set m.IbIdsByAccount = Nothing
    Set m.TradeRoutes = Nothing
    Set m.CarPosFills = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRithmic.Class_Terminate"
    
End Sub

