VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerCqg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerCqg.cls
'' Description: Class to manange communications with CQG servers
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/25/2011   DAJ         Build current position, use set for objects back from BrokerKv
'' 09/07/2011   DAJ         Various fixes including amending orders
'' 09/26/2011   DAJ         Fix for calculating position for symbol with no carried position
'' 10/25/2011   DAJ         Pass broker object to build position call
'' 11/02/2011   DAJ         Added Amp Trading and RJ O'Brien as CQG brokers
'' 11/18/2011   DAJ         Fix for converting future option symbols
'' 12/06/2011   DAJ         Changed "Broker Name" for RJ O'Brien
'' 12/13/2011   DAJ         Added Capital Trading Group for CQG
'' 01/30/2012   DAJ         Allow Reconnect flag on Connection Status message
'' 03/14/2012   DAJ         Added Penson(CQG)
'' 04/02/2012   DAJ         Fix for handling multiple positions without carried info
'' 04/05/2012   DAJ         Added broker mode code
'' 05/01/2012   DAJ         Added account details calls
'' 05/31/2012   DAJ         Turnkey implementation
'' 06/06/2012   DAJ         Added the subsender to the login call
'' 06/11/2012   DAJ         Make Turnkey work with all brokers
'' 06/12/2012   DAJ         Make sure that account doesn't get created in broker mode
'' 07/16/2012   DAJ         ZanerCqg, KnightCqg
'' 07/17/2012   DAJ         RobbinsCqg
'' 07/31/2012   DAJ         If Carried Position is null, set it to zero
'' 08/28/2012   DAJ         CarriedFillFromInfo, First Field Begin/End
'' 10/05/2012   DAJ         Added Account threshold for CQG brokers
'' 10/09/2012   DAJ         Set the PreviousOrderID in the OrderFromInfo routine if applicable
'' 12/11/2012   DAJ         Broker enabled symbols for trading
'' 12/11/2012   DAJ         Vision (CQG)
'' 01/11/2013   DAJ         Added symbol conversion code for future calendar spreads
'' 01/18/2013   DAJ         CQG symbols enabled for market data
'' 01/29/2013   DAJ         When adding carried fills to position, pay attention to account
'' 03/06/2013   DAJ         For L/F, ignore enabled symbols check if list is empty
'' 03/11/2013   DAJ         Handle "Reject Message" from CQG, Last known symbols
'' 04/02/2013   DAJ         Don't do account filter on Get's if not a broker
'' 04/03/2013   DAJ         Log when removing order from amend orders array
'' 04/04/2013   DAJ         Added GF and LE to the default broker override list
'' 04/09/2013   DAJ         Added the GetPositionForFill routine
'' 04/16/2013   DAJ         Call 'PendingNew' orders 'Queued' instead of 'PreSubmitted'
'' 04/17/2013   DAJ         Make sure that the enabled symbols list is sorted, give LF real-time
'' 05/10/2013   DAJ         If an order comes back 'cancelled' with 'reason#77', call it expired
'' 05/21/2013   DAJ         In the GenesisSymbolList routine, always pass back all symbols if LF
'' 10/16/2013   DAJ         Add FptCqg
'' 01/09/2014   DAJ         Changed BrokerName for FuturePath Trading
'' 03/07/2014   DAJ         Moved Cattle stuff into NavCattle.DLL
'' 06/03/2014   DAJ         Hook up amended order that came in as "Partial" to the old order
'' 09/09/2014   DAJ         Handle 'Cancelled: 'marketClose'' message as expired order
'' 11/14/2014   DAJ         Added support for MIT orders
'' 04/13/2015   DAJ         Don't show "Order updated at exchange..." message as rejected order message
'' 04/28/2015   DAJ         Change 'Suspended' order status to 'PreSubmitted'
'' 02/17/2016   DAJ         Handle 'Order is expired' message as expired order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kRegKey = "Software\Genesis Financial Data Services\Navigator Suite\General"

Private Type mPrivate
    Broker As cBroker                   ' Broker base class object
    BrokerKv As cBrokerKeyValue         ' Routines for key-value broker messages
    Symbols As cBrokerSymbols           ' Broker symbol conversion object
    
    strUserName As String               ' User name to login to the server with
    strPassword As String               ' Password for logging into the server
    strIP As String                     ' IP address for the server
    strPort As String                   ' Port for the server
    strTargetID As String               ' Target computer ID
    strSubsender As String              ' Subsender ID
    
    AmendOrders As cGdTree              ' Collection of orders that have been sent for amend
    astrFillSymbols As cGdArray         ' Array of symbols for which we got fills
    astrEnabledSymbols As cGdArray      ' Array of Genesis symbols the user is enabled for
    
    Accounts As cGdTree                 ' Collection of accounts received
    Orders As cGdTree                   ' Collection of orders received
    Fills As cGdTree                    ' Collection of fills received
End Type
Private m As mPrivate

Public Property Get Broker() As cBroker
    Set Broker = m.Broker
End Property

Public Property Get AmendOrders() As cGdTree
    Set AmendOrders = m.AmendOrders
End Property

Public Property Get Accounts() As cGdTree
    Set Accounts = m.Accounts
End Property

Public Property Get Orders() As cGdTree
    Set Orders = m.Orders
End Property

Public Property Get Fills() As cGdTree
    Set Fills = m.Fills
End Property

Public Property Get EnabledSymbols() As cGdArray
    If (IsLeaderFollower = True) And (m.astrEnabledSymbols.Size = 0) Then
        Set EnabledSymbols = Nothing
    Else
        Set EnabledSymbols = m.astrEnabledSymbols
    End If
End Property

Private Property Get IsLeaderFollower() As Boolean
    IsLeaderFollower = ((UCase(m.strUserName) = "LF_APP") Or (UCase(m.strTargetID) = "LF_GATEWAY"))
End Property

Private Property Get LKSCRC() As Long
    LKSCRC = GetRegistryValue(rkLocalMachine, kRegKey, "LKSC" & Str(m.Broker.Broker), 0&)
End Property
Private Property Let LKSCRC(ByVal lLksCrc As Long)
    SetRegistryValue rkLocalMachine, kRegKey, "LKSC" & Str(m.Broker.Broker), lLksCrc, True
End Property

Private Property Get LksFile() As String
    LksFile = AddSlash(App.Path) & "Provided\LKS." & Str(m.Broker.Broker)
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the object based on what broker it is
'' Inputs:      Broker, Is user a broker?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal nBroker As eTT_AccountType, Optional ByVal bIsBroker As Boolean = False)
On Error GoTo ErrSection:

    Dim strSymbolFile As String         ' Symbol conversion file
    Dim strBrokerPath As String         ' Broker path
    
    Select Case nBroker
        Case eTT_AccountType_CQG
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "CQG", True
            m.Broker.IniFile = AddSlash(App.Path) & "CQG.INI"
            m.Broker.BrokerName = "CQG"
            m.Broker.ProcessName = "GenCQG"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\Cqg\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenCqg.EXE"
            m.Broker.ControlID = "CQG"
            strSymbolFile = AddSlash(App.Path) & "Provided\CqgToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\CqgCnct.INI"
            
        Case eTT_AccountType_AmpCqg
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "AmpCqg", True
            m.Broker.IniFile = AddSlash(App.Path) & "AmpCqg.INI"
            m.Broker.BrokerName = "AMP"
            m.Broker.ProcessName = "GenAmpCqg"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\AmpCqg\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenAmpCqg.EXE"
            m.Broker.ControlID = "AMPCQG"
            strSymbolFile = AddSlash(App.Path) & "Provided\AmpCqgToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\AmpCqgCnct.INI"
            
        Case eTT_AccountType_CtgCqg
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "CtgCqg", True
            m.Broker.IniFile = AddSlash(App.Path) & "CtgCqg.INI"
            m.Broker.BrokerName = "CTG2"
            m.Broker.ProcessName = "GenCtgCqg"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\CtgCqg\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenCtgCqg.EXE"
            m.Broker.ControlID = "CtgCQG"
            strSymbolFile = AddSlash(App.Path) & "Provided\CtgCqgToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\CtgCqgCnct.INI"
            
        Case eTT_AccountType_FptCqg
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "FptCqg", True
            m.Broker.IniFile = AddSlash(App.Path) & "FptCqg.INI"
            m.Broker.BrokerName = "FuturePath Trading (CQG)"
            m.Broker.ProcessName = "GenFptCqg"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\FptCqg\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenFptCqg.EXE"
            m.Broker.ControlID = "FptCQG"
            strSymbolFile = AddSlash(App.Path) & "Provided\FptCqgToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\FptCqgCnct.INI"
            
        Case eTT_AccountType_KnightCqg
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "KnightCqg", True
            m.Broker.IniFile = AddSlash(App.Path) & "KnightCqg.INI"
            m.Broker.BrokerName = "Knight (CQG)"
            m.Broker.ProcessName = "GenKnightCqg"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\KnightCqg\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenKnightCqg.EXE"
            m.Broker.ControlID = "KnightCQG"
            strSymbolFile = AddSlash(App.Path) & "Provided\KnightCqgToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\KnightCqgCnct.INI"
                
        Case eTT_AccountType_RjoCqg
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "RjoCqg", True
            m.Broker.IniFile = AddSlash(App.Path) & "RjoCqg.INI"
            m.Broker.BrokerName = "RJ O'Brien (CQG)"
            m.Broker.ProcessName = "GenRjoCqg"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\RjoCqg\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenRjoCqg.EXE"
            m.Broker.ControlID = "RJOCQG"
            strSymbolFile = AddSlash(App.Path) & "Provided\RjoCqgToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\RjoCqgCnct.INI"
            
        Case eTT_AccountType_RobbinsCqg
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "RobbinsCqg", True
            m.Broker.IniFile = AddSlash(App.Path) & "RobbinsCqg.INI"
            m.Broker.BrokerName = "Robbins (CQG)"
            m.Broker.ProcessName = "GenRobbinsCqg"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\RobbinsCqg\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenRobbinsCqg.EXE"
            m.Broker.ControlID = "RobbinsCQG"
            strSymbolFile = AddSlash(App.Path) & "Provided\RobbinsCqgToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\RobbinsCqgCnct.INI"
            
        Case eTT_AccountType_VisionCqg
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "VisionCqg", True
            m.Broker.IniFile = AddSlash(App.Path) & "VisionCqg.INI"
            m.Broker.BrokerName = "Vision (CQG)"
            m.Broker.ProcessName = "GenVisionCqg"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\VisionCqg\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenVisionCqg.EXE"
            m.Broker.ControlID = "VisionCQG"
            strSymbolFile = AddSlash(App.Path) & "Provided\VisionCqgToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\VisionCqgCnct.INI"
            
        Case eTT_AccountType_ZanerCqg
            m.Broker.Init Me, nBroker, AddSlash(App.Path) & "ZanerCqg", True
            m.Broker.IniFile = AddSlash(App.Path) & "ZanerCqg.INI"
            m.Broker.BrokerName = "Zaner (CQG)"
            m.Broker.ProcessName = "GenZanerCqg"
            strBrokerPath = AddSlash(App.Path) & "..\Brokers\ZanerCqg\"
            m.Broker.ProcessPath = AddSlash(strBrokerPath) & "GenZanerCqg.EXE"
            m.Broker.ControlID = "ZanerCQG"
            strSymbolFile = AddSlash(App.Path) & "Provided\ZanerCqgToGen.TXT"
            m.Broker.ConnectIni = AddSlash(App.Path) & "Provided\ZanerCqgCnct.INI"
            
    End Select
    
    m.Broker.IsBroker = bIsBroker

    m.Broker.CopyProcess AddSlash(App.Path) & "..\Brokers\GenCqg.EXE"
    FileCopy AddSlash(App.Path) & "..\Brokers\FixDefCqg.xml", AddSlash(strBrokerPath) & "FixDef.xml", True
    
    Set m.Symbols = New cBrokerSymbols
    m.Symbols.FromFile strSymbolFile
    
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Connect) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Disconnect) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_SwitchAccounts) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_SwitchAccountsMode) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ConnectInfo) = True
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ChangePassword) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_Refresh) = Not bIsBroker
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewActivity) = Not bIsBroker
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_BrokerView) = bIsBroker
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_ViewOnline) = False
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_VerifyPositions) = Not bIsBroker
    m.Broker.ShowTradeMenuItem(eGDTradingMenu_AccountDetails) = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.Init"

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

    ShowConfig = ((Len(m.strUserName) = 0) Or (Len(m.strPassword) = 0) Or (Len(m.strIP) = 0) Or (Len(m.strPort) = 0) Or (Len(m.strTargetID) = 0))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.ShowConfig"
    
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

    ShowBrokerConnectionInfo = GetLoginInfo(strUserName, , True)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.ShowBrokerConnectionInfo"
    
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
        If (m.Broker.AskedPassword = True) And ((strNewUserName = m.strUserName) Or (Len(strNewUserName) = 0)) Then
            bReturn = True
        Else
            bReturn = GetLoginInfo(strNewUserName)
        End If
    Else
        If ((strNewUserName <> m.strUserName) And (Len(strNewUserName) > 0)) Or (bSwitching = True) Then
            bReturn = GetLoginInfo(strNewUserName, bSwitching)
            If bReturn = True Then
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
    RaiseError "cBrokerCqg.ShowLogin"
    
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

    m.Broker.DumpDebug "Connecting to " & m.Broker.BrokerName & " as '" & m.strUserName & "'"
    m.Broker.SendBrokerMessage eGDBrokerMessageType_Connect, ConnectString, , ConnectString(True)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.SendConnect"
    
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

    m.Broker.SendBrokerMessage eGDBrokerMessageType_Disconnect, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.SendDisconnect"
    
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

    m.Broker.SendBrokerMessage eGDBrokerMessageType_UnloadApp, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.UnloadApp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerSymbol
'' Description: Attempt to convert the Genesis symbol to the broker symbology
'' Inputs:      Genesis Symbol, Broker Symbol object
'' Returns:     Broker Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BrokerSymbol(ByVal strGenesisSymbol As String, Optional BrokerSym As cBrokerSymbol) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrSymbol As cGdArray          ' Symbol broken into pieces
    Dim OptionSym As cOptionSymbol      ' Option symbol object
    
    strReturn = ""
    If Len(strGenesisSymbol) > 0 Then
        Select Case UCase(SecurityType(strGenesisSymbol, True))
            Case "F"
                Set astrSymbol = New cGdArray
                astrSymbol.SplitFields strGenesisSymbol, "-"
                
                If astrSymbol.Size = 2 Then
                    Set BrokerSym = m.Symbols.GetByGenesisSymbol(astrSymbol(0))
                    If Not BrokerSym Is Nothing Then
                        strReturn = "F.US." & BrokerSym.BrokerBase & BrokerSym.ContractToMYY(astrSymbol(1))
                    End If
                ElseIf astrSymbol.Size = 3 Then
                    Set BrokerSym = m.Symbols.GetByGenesisSymbol("S:" & astrSymbol(0))
                    If Not BrokerSym Is Nothing Then
                        strReturn = "F.US." & Replace(BrokerSym.BrokerBase, "S:", "") & Right(astrSymbol(2), Len(astrSymbol(2)) - 1) & BrokerSym.ContractToMYY(astrSymbol(1))
                    End If
                End If
                
            Case "FO"
                Set OptionSym = New cOptionSymbol
                OptionSym.FromGenesis strGenesisSymbol
                Set BrokerSym = m.Symbols.GetByGenesisSymbol("O:" & OptionSym.BaseSymbol)
                If Not BrokerSym Is Nothing Then
                    strReturn = OptionSym.ToCqg(Replace(BrokerSym.BrokerBase, "O:", ""), BrokerSym.StrikeMult)
                End If
        
        End Select
    End If
    
    BrokerSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.BrokerSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbol
'' Description: Attempt to convert the broker symbol to the Genesis symbology
'' Inputs:      Broker Symbol, Broker Exchange, Broker Symbol object, For Position?
'' Returns:     Genesis Symbol (Blank if cannot convert)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbol(ByVal strBrokerSymbol As String, ByVal strBrokerExchange As String, Optional BrokerSym As cBrokerSymbol, Optional ByVal bPosition As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrSymbol As cGdArray          ' Symbol broken up into fields
    Dim OptionSym As cOptionSymbol      ' Option symbol object
    Dim strBase As String               ' Base portion of the symbol
    Dim strContract As String           ' Contract portion of the symbol
    Dim strLegGap As String             ' CQG strategy leg gap
    Dim strGenesisBaseSymbol As String  ' Genesis base symbol
    Dim lIndexLastNonDigit As Long      ' Index in the symbol of the last non-digit
    
    strReturn = ""
    Set astrSymbol = New cGdArray
    astrSymbol.SplitFields strBrokerSymbol, "."
    
    If astrSymbol.Size = 3 Then
        lIndexLastNonDigit = IndexOfLastNonDigit(strBrokerSymbol)
        If lIndexLastNonDigit >= 0 Then
            strBase = Left(strBrokerSymbol, lIndexLastNonDigit - 1)
            strGenesisBaseSymbol = GenesisBaseSymbol(strBase, bPosition, strLegGap, BrokerSym)
            If Len(strGenesisBaseSymbol) > 0 Then
                If UCase(astrSymbol(0)) = "F" Then
                    ' Future: F.US.GLEG13
                    ' Spread: F.US.GLES1G13
                    strContract = Right(astrSymbol(2), 3)
                    
                    If InStr(strGenesisBaseSymbol, "S:") = 0 Then
                        strReturn = strGenesisBaseSymbol & "-" & BrokerSym.ContractFromMYY(strContract)
                    Else
                        strReturn = Replace(strGenesisBaseSymbol, "S:", "") & "-" & BrokerSym.ContractFromMYY(strContract) & "-S" & strLegGap
                    End If
                    
                ElseIf (UCase(astrSymbol(0)) = "C") Or (UCase(astrSymbol(0)) = "P") Then
                    ' {C|P}.US.<BaseSym><M><YY><Strike> (e.g. C.US.QOJ091355)
                    Set OptionSym = New cOptionSymbol
                    OptionSym.FromCqg strBrokerSymbol
                    OptionSym.ConvertBrokerInfo Replace(strGenesisBaseSymbol, "O:", ""), BrokerSym.StrikeMult
                    strReturn = OptionSym.ToGenesis
                End If
            End If
        End If
    End If
    
    GenesisSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.GenesisSymbol"
    
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

    BrokerOrderType = m.BrokerKv.BrokerOrderType(nGenesisOrderType)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.BrokerOrderType"
    
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

    Dim strOrderString As String        ' Broker order string
    
    strOrderString = BrokerOrderString(Order, True)
    
    AddOrder = AddOrderFromString(strOrderString)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.AddOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddOrderFromString
'' Description: Attempt to add the order
'' Inputs:      String to send to broker
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AddOrderFromString(ByVal strOrderString As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If (Len(strOrderString) > 0) Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_AddOrder, strOrderString
        bReturn = True
    End If
    
    AddOrderFromString = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.AddOrderFromString"
    
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

    Dim strOrderString As String        ' Broker order string
    Dim NewOrder As cPtOrder            ' New Order object
    
    Set NewOrder = Order.MakeCopy
    NewOrder.GenesisOrderID = NextGenesisID(g.Broker.AccountNumberForID(Order.AccountID))
    strOrderString = BrokerOrderString(NewOrder, False, Order.GenesisOrderID)
        
    AmendOrder = AmendOrderFromString(strOrderString)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.AmendOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AmendOrderFromString
'' Description: Attempt to amend the order
'' Inputs:      String to send to broker
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AmendOrderFromString(ByVal strOrderString As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If (Len(strOrderString) > 0) Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_AmendOrder, strOrderString
        bReturn = True
    End If
    
    AmendOrderFromString = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.AmendOrderFromString"
    
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

    Dim strOrderString As String        ' Broker order string
    Dim NewOrder As cPtOrder            ' New Order object
    
    Set NewOrder = Order.MakeCopy
    NewOrder.GenesisOrderID = NextGenesisID(g.Broker.AccountNumberForID(Order.AccountID))
    strOrderString = BrokerOrderString(NewOrder, False, Order.GenesisOrderID)
    
    CancelOrder = CancelOrderFromString(strOrderString)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.CancelOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CancelOrderFromString
'' Description: Attempt to cancel the order
'' Inputs:      String to send to broker
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CancelOrderFromString(ByVal strOrderString As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If (Len(strOrderString) > 0) Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_CancelOrder, strOrderString
        bReturn = True
    End If
    
    CancelOrderFromString = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.CancelOrderFromString"
    
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
    Dim NewOrder As cPtOrder            ' New Order object
    
    bReturn = False
    
    Set NewOrder = Order.MakeCopy
    NewOrder.GenesisOrderID = NextGenesisID(g.Broker.AccountNumberForID(Order.AccountID))
    strOrderString = BrokerOrderString(NewOrder, False, Order.GenesisOrderID)
        
    If (Len(strOrderString) > 0) Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_CancelOrder, strOrderString
        bReturn = True
    End If
    
    ParkOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.ParkOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrdersAsOco
'' Description: Submit the given orders as an Order-Cancels-Order
'' Inputs:      Order1, Order2
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SubmitOrdersAsOco(Order1 As cPtOrder, Order2 As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strOrderString1 As String       ' Broker order string for Order 1
    Dim strOrderString2 As String       ' Broker order string for Order 2
    
    bReturn = False
    strOrderString1 = BrokerOrderString(Order1, True)
    strOrderString2 = BrokerOrderString(Order2, True)
    
    If (Len(strOrderString1) > 0) And (Len(strOrderString2) > 0) Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_AddOcoOrders, "BEGIN"
        m.Broker.SendBrokerMessage eGDBrokerMessageType_AddOcoOrders, strOrderString1
        m.Broker.SendBrokerMessage eGDBrokerMessageType_AddOcoOrders, strOrderString2
        m.Broker.SendBrokerMessage eGDBrokerMessageType_AddOcoOrders, "END"
        
        bReturn = True
    End If

    SubmitOrdersAsOco = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.SubmitOrdersAsOco"
    
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
    RaiseError "cBrokerCqg.Refresh"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetNumberOfAccounts
'' Description: Request number of accounts from the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetNumberOfAccounts()
On Error GoTo ErrSection:

    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetNumberOfAccounts, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.GetNumberOfAccounts"
    
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

    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetAccounts, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.GetAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetAccountDetails
'' Description: Request account details from the server
'' Inputs:      Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAccountDetails(ByVal strAccount As String)
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker message
    
    Set brokerMsg = New cBrokerMessage
    brokerMsg.Add "Account", strAccount
    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetAccountDetails, brokerMsg.ToString

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.GetAccountDetails"
    
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

    Dim brokerMsg As cBrokerMessage     ' Broker message
    
    If (Len(strAccount) = 0) Or (m.Broker.IsBroker = False) Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_GetOrders, ""
    Else
        Set brokerMsg = New cBrokerMessage
        brokerMsg.Add "Account", strAccount
        m.Broker.SendBrokerMessage eGDBrokerMessageType_GetOrders, brokerMsg.ToString
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.GetOrders"
    
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

    Dim brokerMsg As cBrokerMessage     ' Broker message
    
    If (Len(strAccount) = 0) Or (m.Broker.IsBroker = False) Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_GetFills, ""
    Else
        Set brokerMsg = New cBrokerMessage
        brokerMsg.Add "Account", strAccount
        m.Broker.SendBrokerMessage eGDBrokerMessageType_GetFills, brokerMsg.ToString
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.GetFills"
    
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

    Dim brokerMsg As cBrokerMessage     ' Broker message
    
    If (Len(strAccount) = 0) Or (m.Broker.IsBroker = False) Then
        m.Broker.SendBrokerMessage eGDBrokerMessageType_GetPositions, ""
    Else
        Set brokerMsg = New cBrokerMessage
        brokerMsg.Add "Account", strAccount
        m.Broker.SendBrokerMessage eGDBrokerMessageType_GetPositions, brokerMsg.ToString
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.GetPositions"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSymbols
'' Description: Request symbols from the server
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetSymbols()
On Error GoTo ErrSection:

    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetSymbols, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.GetSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSides
'' Description: Request a list of valid side values
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetSides()
On Error GoTo ErrSection:
    
    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetSides, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.GetSides"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetTifs
'' Description: Request a list of valid time-in-force values
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetTifs()
On Error GoTo ErrSection:
    
    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetTifs, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.GetTifs"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrderTypes
'' Description: Request a list of valid order type values
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetOrderTypes()
On Error GoTo ErrSection:
    
    m.Broker.SendBrokerMessage eGDBrokerMessageType_GetOrderTypes, ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.GetOrderTypes"
    
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
    Dim strMachineID As String          ' Machine ID
    Dim strIniFile As String            ' Ini File
    
    strIniFile = m.Broker.IniFile
    lDate = GetIniFileProperty("Date", 0&, "ID", strIniFile)
    lCounter = GetIniFileProperty("Counter", 0&, "ID", strIniFile)
    strMachineID = StripStr(UCase(RI_GetMachineID), "- ")
    
    If lDate <> Date Then
        lDate = Date
        lCounter = 0&
    Else
        lCounter = lCounter + 1
    End If

    SetIniFileProperty "Date", lDate, "ID", strIniFile
    SetIniFileProperty "Counter", lCounter, "ID", strIniFile
    
    NextGenesisID = "TN_" & strMachineID & "_" & Format(lDate, "00000") & Format(lCounter, "00000")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.NextGenesisID"
    
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
    RaiseError "cBrokerCqg.StartingGenesisIdForOptNav"
    
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
    RaiseError "cBrokerCqg.IsTradeableSymbol"
    
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
    Dim strSecurityType As String       ' Security type for the symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    
    bReturn = (Len(BrokerSymbol(strGenesisSymbol, BrokerSym)) > 0)
    If bReturn Then
        strBrokerBase = BrokerSym.BrokerBase
        strBrokerExchange = BrokerSym.BrokerExchange
        
        ' DAJ 04/18/2013: If they are a leader/follower user, just give them everything since
        ' we know they have a live account...
        'If (IsLeaderFollower = True) And (m.astrEnabledSymbols.Size = 0) Then
        If (IsLeaderFollower = True) Then
            bReturn = True
        Else
            strSecurityType = SecurityType(strGenesisSymbol, True)
            
            Select Case UCase(strSecurityType)
                Case "F"
                    bReturn = m.astrEnabledSymbols.BinarySearch(Parse(strGenesisSymbol, "-", 1))
                
                Case "FO"
                    bReturn = m.astrEnabledSymbols.BinarySearch("O:" & Parse(strGenesisSymbol, "-", 1))
                
                Case Else
                    bReturn = m.astrEnabledSymbols.BinarySearch(strGenesisSymbol)
                    
            End Select
        End If
    End If
    
    IsEnabledSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.IsEnabledSymbol"
    
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

    Dim bReturn As Boolean              ' Return value from the function
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    
    bReturn = False
    If Len(BrokerSymbol(strGenesisSymbol, BrokerSym)) > 0 Then
        bReturn = BrokerSym.OrderTypeAllowed(nOrderType)
    End If
    
    OrderTypeAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.OrderTypeAllowed"
    
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

    Dim bReturn As Boolean              ' Return value from the function
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    
    bReturn = False
    If Len(BrokerSymbol(strGenesisSymbol, BrokerSym)) > 0 Then
        bReturn = BrokerSym.TifAllowed(nTimeInForce)
    End If
    
    TimeInForceAllowed = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.TimeInForceAllowed"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleMessage
'' Description: Handle an incoming App Mail message from the stand-alone program
'' Inputs:      Message Type, Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleMessage(ByVal nType As eGDBrokerMessageTypes, ByVal strMessage As String)
On Error GoTo ErrSection:
    
    Select Case nType
        Case eGDBrokerMessageType_AccountDetails
            If (m.Broker.IsBroker = True) And (FormIsLoaded("frmBrokerView") = True) Then
                frmBrokerView.Broker_AccountDetails strMessage
            End If
    
        Case eGDBrokerMessageType_OrderTypes
            frmTest2.AddList "Order Types = '" & strMessage & "'"
        
        Case eGDBrokerMessageType_Sides
            frmTest2.AddList "Sides = '" & strMessage & "'"
            
        Case eGDBrokerMessageType_Symbols
            HandleSymbols strMessage
        
        Case eGDBrokerMessageType_Tifs
            frmTest2.AddList "Tifs = '" & strMessage & "'"
            
        Case eGDBrokerMessageType_RejectMessage
            HandleRejectMessage strMessage
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerTemplate.HandleMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    StatusFromInfo
'' Description: Get connection status information from the given message
'' Inputs:      Message, Status, Error, User Name, Allow Reconnect?
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StatusFromInfo(ByVal strMessage As String, nStatus As eGDConnectionStatus, strError As String, strUserName As String, bAllowReconnect As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    
    Set brokerMsg = New cBrokerMessage
    brokerMsg.FromString strMessage
    
    nStatus = CLng(Val(brokerMsg("Status")))
    strError = brokerMsg("Error")
    strUserName = m.strUserName
    
    If Len(brokerMsg("AllowReconnect")) = 0 Then
        bAllowReconnect = True
    Else
        bAllowReconnect = (CLng(Val(brokerMsg("AllowReconnect"))) <> 0)
    End If
    
    StatusFromInfo = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.StatusFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountFromInfo
'' Description: Fill an account object from the given message
'' Inputs:      Message, Account
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountFromInfo(ByVal strMessage As String, Acct As cPtAccount) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
            m.Accounts.Clear
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            brokerMsg.FromString strMessage
            
            Set Acct = m.BrokerKv.AccountFromMessage(brokerMsg, m.Broker, m.strUserName, m.strPassword)
            
            If m.Accounts.Exists(brokerMsg("Account")) Then
                Set m.Accounts(brokerMsg("Account")) = brokerMsg
            Else
                m.Accounts.Add brokerMsg, brokerMsg("Account")
            End If
            
            bReturn = True
        End If
    End If
    
    AccountFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.AccountFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderFromInfo
'' Description: Fill an order object from the given message
'' Inputs:      Message, Order
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderFromInfo(ByVal strMessage As String, Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim brokerMsg As cBrokerMessage     ' Broker message object
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim lQuantityFilled As Long         ' Quantity of the order that has filled
    Dim Bars As New cGdBars             ' Bars object
    Dim AmendedOrder As cPtOrder        ' Order object
    Dim strPrevGenesisID As String      ' Previous Genesis ID
    Dim strError As String              ' Error Message
    Dim Account As cBrokerMessage       ' Account
    Dim lAccountID As Long              ' Account ID
    Dim strFirstField As String         ' First field in the string
    Dim lPreviousOrderID As Long        ' Previous Order ID

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
            m.Orders.Clear
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            lPreviousOrderID = 0&
            
            brokerMsg.FromString strMessage
            
            strGenesisSymbol = GenesisSymbol(brokerMsg("Symbol"), "", BrokerSym)
            strError = brokerMsg("Message")
            strPrevGenesisID = brokerMsg("PreviousGenesisID")
            
            If Len(strPrevGenesisID) > 0 Then
                ' 06/03/2014 DAJ: Customer Tom Williams ran into an issue where he sent an order amendment
                ' on an order while a partial fill was coming in.  When the new order went working, it came
                ' in as "Partial" instead of "Replaced" or "New" and we didn't "hook up" the old order
                ' correctly.  So, I am including "Partial" here...
                If (UCase(brokerMsg("Status")) = "REPLACED") Or (UCase(brokerMsg("Status")) = "NEW") Or (UCase(brokerMsg("Status")) = "PARTIAL") Then
                    If m.AmendOrders.Exists(strPrevGenesisID) Then
                        If Len(strError) > 0 Then
                            Set AmendedOrder = m.AmendOrders(strPrevGenesisID)
                            AmendedOrder.ChangeOrderStatus eTT_OrderStatus_Working
                            
                            InfBox strError, , , m.Broker.BrokerName & " Order Error"
                            strGenesisSymbol = ""
                        Else
                            Set AmendedOrder = m.AmendOrders(strPrevGenesisID)
                            AmendedOrder.ChangeOrderStatus eTT_OrderStatus_Amended
                            lPreviousOrderID = AmendedOrder.OrderID
                        End If
                        
                        m.Broker.DumpDebug vbTab & "Order with Broker ID '" & strPrevGenesisID & "' removed from amend orders array"
                        m.AmendOrders.Remove strPrevGenesisID
                    End If
                    
                    If m.Orders.Exists("PreviousBrokerID") Then
                        m.Orders.Remove "PreviousBrokerID"
                    End If
                End If
            End If

            If Len(strGenesisSymbol) > 0 Then
                If m.Broker.IsBroker Then
                    lAccountID = -1&
                Else
                    lAccountID = kNullData
                End If
                
                Set Order = m.BrokerKv.OrderFromMessage(brokerMsg, m.Broker, strGenesisSymbol, BrokerSym, lAccountID)
                
                ' DAJ 04/16/2013: The PendingNew status that comes back from CQG is still considered
                ' to be a pending status.  We don't want to use PreSubmitted because it is considered
                ' to be a "working" status and that it can then be modified, cancelled, etc....
                If Order.Status = eTT_OrderStatus_PreSubmitted Then
                    Order.Status = eTT_OrderStatus_Queued
                End If
                
                ' DAJ 05/10/2013: We discovered that an order that comes back with a status of
                ' 'Cancelled' and a message of 'reason#77' when it expires, so just set the status
                ' to Expired here, but clear the message so that we don't display it to the user...
                ' DAJ 09/09/2014: Discovered that with the softs, we are getting the message
                ' 'Cancelled: 'marketClose'' coming back on an expired message...
                ' DAJ 02/17/2016: Tim started getting the message 'Order is expired", so we will
                ' handle that here...
                If (Order.Status = eTT_OrderStatus_Cancelled) And ((Trim(Order.Message) = "reason#77") Or (Trim(Order.Message) = "Cancelled: 'marketClose'") Or (InStr(UCase(Order.Message), "ORDER IS EXPIRED") <> 0)) Then
                    Order.Message = ""
                    Order.Status = eTT_OrderStatus_Expired
                End If
                
                ' DAJ 04/13/2015: Tim noticed that when a Stop with Limit order changes from a stop
                ' to a limit order, he is getting a message saying "Order updated at exchange..." and
                ' we are displaying it as if it is a rejected order message.  We should just eat the
                ' message in this case...
                If (Order.Status = eTT_OrderStatus_Working) And (InStr(UCase(Order.Message), "ORDER UPDATED AT EXCHANGE") > 0) Then
                    Order.Message = ""
                End If
                
                ' DAJ 04/28/2015: Tim doesn't like the status of 'Suspended' for an order submitted
                ' while the market is closed so we are going to go with 'PreSubmitted' and see what happens...
                If Order.Status = eTT_OrderStatus_Suspended Then
                    Order.Status = eTT_OrderStatus_PreSubmitted
                End If
                
                Order.PreviousOrderID = lPreviousOrderID
                
                brokerMsg.Add "GenesisSymbol", strGenesisSymbol
                If Len(brokerMsg("LimitPrice")) > 0 Then
                    brokerMsg.Add "GenesisLimitPrice", Str(Order.LimitPrice)
                End If
                If Len(brokerMsg("StopPrice")) > 0 Then
                    brokerMsg.Add "GenesisStopPrice", Str(Order.StopPrice)
                End If
                brokerMsg.Add "GenesisType", Str(Order.OrderType)
                brokerMsg.Add "GenesisTif", Str(Order.TimeInForce)
                brokerMsg.Add "GenesisStatus", Str(Order.Status)
                
                If m.Accounts.Exists(brokerMsg("Account")) Then
                    Set Account = m.Accounts(brokerMsg("Account"))
                    If Len(Account("FcmAccount")) > 0 Then
                        brokerMsg.Add "FcmAccount", Account("FcmAccount")
                    Else
                        brokerMsg.Add "FcmAccount", Account("Account")
                    End If
                End If
                
                If IsOpenOrder(Order.Status) Then
                    If m.Orders.Exists(brokerMsg("BrokerID")) Then
                        Set m.Orders(brokerMsg("BrokerID")) = brokerMsg
                    Else
                        m.Orders.Add brokerMsg, brokerMsg("BrokerID")
                    End If
                Else
                    If m.Orders.Exists(brokerMsg("BrokerID")) Then
                        m.Orders.Remove brokerMsg("BrokerID")
                    End If
                End If
                
                g.CattleBridge.Broker_Order g.TnCattle.KeyValueOrderToTurnkey(brokerMsg, m.Broker.Broker)
                
                bReturn = True
            End If
        End If
    End If
    
    OrderFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.OrderFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillFromInfo
'' Description: Fill a fill object from the given message
'' Inputs:      Message, Fill, Genesis Order ID, Carried?
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillFromInfo(ByVal strMessage As String, Fill As cPtFill, Optional strGenesisOrderID As String = "", Optional ByVal bCarried As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim brokerMsg As cBrokerMessage     ' Broker message object
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim Bars As New cGdBars             ' Bars object
    Dim lPos As Long                    ' Position of the symbol in the array
    Dim strKey As String                ' Key into the fills array
    Dim Account As cBrokerMessage       ' Account
    Dim strBrokerID As String           ' Broker Order ID
    Dim lAccountID As Long              ' Account ID
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
            m.astrFillSymbols.Clear
            m.Fills.Clear
        ElseIf UCase(strFirstField) = "END" Then
        Else
            Set brokerMsg = New cBrokerMessage
            
            brokerMsg.FromString strMessage
            
            strGenesisSymbol = GenesisSymbol(brokerMsg("Symbol"), "", BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
                strKey = brokerMsg("Account") & vbTab & brokerMsg("Symbol")
                If m.astrFillSymbols.BinarySearch(strKey, lPos) = False Then
                    m.astrFillSymbols.Add strKey, lPos
                End If
                
                strBrokerID = brokerMsg("BrokerID")
                brokerMsg("BrokerID") = ""
                If m.Broker.IsBroker Then
                    lAccountID = -1&
                Else
                    lAccountID = kNullData
                End If
                Set Fill = m.BrokerKv.FillFromMessage(brokerMsg, m.Broker, strGenesisSymbol, BrokerSym, lAccountID)
                strGenesisOrderID = brokerMsg("GenesisID")
                
                brokerMsg("BrokerID") = strBrokerID
                brokerMsg.Add "GenesisSymbol", strGenesisSymbol
                brokerMsg.Add "GenesisFillPrice", Str(Fill.Price)
                
                If Len(brokerMsg("FillID")) = 0 Then
                End If
                
                If m.Accounts.Exists(brokerMsg("Account")) Then
                    Set Account = m.Accounts(brokerMsg("Account"))
                    If Len(Account("FcmAccount")) > 0 Then
                        brokerMsg.Add "FcmAccount", Account("FcmAccount")
                    Else
                        brokerMsg.Add "FcmAccount", Account("Account")
                    End If
                End If
                
                If Len(brokerMsg("FillID")) > 0 Then
                    If m.Fills.Exists(brokerMsg("FillID")) Then
                        Set m.Fills(brokerMsg("FillID")) = brokerMsg
                    Else
                        m.Fills.Add brokerMsg, brokerMsg("FillID")
                    End If
                End If
                
                g.CattleBridge.Broker_Fill g.TnCattle.KeyValueFillToTurnkey(brokerMsg, m.Broker.Broker)
                
                bReturn = True
            End If
        End If
    End If
    
    FillFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.FillFromInfo"
    
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
    RaiseError "cBrokerCqg.CarriedFillFromInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionFromInfo
'' Description: Fill a position object from the given message
'' Inputs:      Message, Position
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PositionFromInfo(ByVal strMessage As String, Position As cPtPosition) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim brokerMsg As cBrokerMessage     ' Broker message object
    Dim strGenesisSymbol As String      ' Genesis symbol
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim lIndex As Long                  ' Index into a for loop
    Dim carriedFill As cPtFill          ' Carried fill
    Dim Fills As cPtFills               ' Fills
    Dim lPos As Long                    ' Position of the symbol in the array
    Dim strKey As String                ' Key into the fills array
    Dim lAccountID As Long              ' Account ID
    Dim strFirstField As String         ' First field in the string

    bReturn = False
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If UCase(strFirstField) = "BEGIN" Then
        ElseIf UCase(strFirstField) = "END" Then
            ' 04/02/2012 DAJ: Each element in this array will end up getting removed
            ' when this function is reentered due to the PositionReceived call inside
            ' the loop.  Because of that, we need to walk backward through the array...
            For lIndex = m.astrFillSymbols.Size - 1 To 0 Step -1
                Set brokerMsg = New cBrokerMessage
                brokerMsg.Add "Account", Parse(m.astrFillSymbols(lIndex), vbTab, 1)
                brokerMsg.Add "Symbol", Parse(m.astrFillSymbols(lIndex), vbTab, 2)
                
                m.Broker.PositionReceived brokerMsg.ToString, True
            Next lIndex
            
            GetSymbols
        Else
            Set brokerMsg = New cBrokerMessage
            brokerMsg.FromString strMessage
            
            strGenesisSymbol = GenesisSymbol(brokerMsg("Symbol"), "", BrokerSym)
            If Len(strGenesisSymbol) > 0 Then
                strKey = brokerMsg("Account") & vbTab & brokerMsg("Symbol")
                If m.astrFillSymbols.BinarySearch(strKey, lPos) = True Then
                    m.astrFillSymbols.Remove lPos
                End If
                
                If m.Broker.IsBroker Then
                    lAccountID = -1&
                Else
                    lAccountID = kNullData
                End If
                Set Position = m.BrokerKv.PositionFromMessage(brokerMsg, m.Broker, strGenesisSymbol, lAccountID)
                
                ' DAJ 07/30/2012: Since we do get a carried position from CQG, this should never
                ' be null.  We are running into cases where we come back through for a symbol
                ' that we got fills for but no carried information that it isn't reset to zero...
                If Position.CarriedPosition = kNullData Then
                    Position.CarriedPosition = 0&
                End If
                
                Position.CurrentPosition = 0
                If Not m.Broker Is Nothing Then
                    For lIndex = 1 To m.Broker.CarriedFills.Count
                        Set carriedFill = m.Broker.CarriedFills(lIndex)
                        If carriedFill.Symbol = strGenesisSymbol Then
                            If carriedFill.AccountID = Position.AccountID Then
                                Position.CarriedFills.Add carriedFill
                                If carriedFill.Buy Then
                                    Position.CurrentPosition = Position.CurrentPosition + carriedFill.Quantity
                                Else
                                    Position.CurrentPosition = Position.CurrentPosition - carriedFill.Quantity
                                End If
                            End If
                        End If
                    Next lIndex
                    
                    Set Fills = m.Broker.BrokerInfo.Fills
                    For lIndex = 1 To Fills.Count
                        If Fills(lIndex).AccountID = Position.AccountID Then
                            If Fills(lIndex).SymbolOrSymbolID = Position.SymbolOrSymbolID Then
                                If Fills(lIndex).Buy Then
                                    Position.CurrentPosition = Position.CurrentPosition + Fills(lIndex).Quantity
                                Else
                                    Position.CurrentPosition = Position.CurrentPosition - Fills(lIndex).Quantity
                                End If
                            End If
                        End If
                    Next lIndex
                End If
                
                bReturn = True
            End If
        End If
    End If
    
    PositionFromInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.PositionFromInfo"
    
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

    Dim strFileName As String           ' Filename for the LKSyms file
    Dim lLksCrc As Long                 ' CRC for the LKSyms file

    If IsLeaderFollower Then
        Set m.astrEnabledSymbols = m.Symbols.GenesisSymbolList(bOnlyIfStream)
    
    ElseIf m.astrEnabledSymbols.Size = 0 Then
        strFileName = LksFile
        lLksCrc = LKSCRC

        If CalcFileCrc(strFileName) <> LKSCRC Then
            KillFile strFileName
        End If

        If m.astrEnabledSymbols.Serialize(strFileName, False) = False Then
            ' If there is no CRC entry in the registry and they have connected before, give
            ' them all of the symbols like we have been doing...
            If (LKSCRC = 0&) And (m.Broker.LastDateConnected > 0) Then
                Set m.astrEnabledSymbols = m.Symbols.GenesisSymbolList(bOnlyIfStream)
            
            ' Otherwise, if the CRC's don't match, or they don't have a LKS file, give them
            ' a small default list...
            Else
                m.astrEnabledSymbols.SplitFields "G6E,ES,GF,LE,NQ,YM,ZB,ZC,ZN", ","
            End If
            
            m.astrEnabledSymbols.Sort
            m.astrEnabledSymbols.Serialize strFileName, True
            LKSCRC = CalcFileCrc(strFileName)
        Else
            m.astrEnabledSymbols.Sort
        End If
    End If

    Set GenesisSymbolList = m.astrEnabledSymbols

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.GenesisSymbolList"
    
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
    RaiseError "cBrokerCqg.TimeZone"
    
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

    'GetAccounts
    GetNumberOfAccounts

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.JustConnected"
    
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
    RaiseError "cBrokerCqg.SymbolInformation"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPositionForFill
'' Description: Generate a position request when we get an unsolicted fill
'' Inputs:      Fill Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPositionForFill(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim FillMessage As cBrokerMessage   ' Broker message object for the fill
    Dim strFirstField As String         ' First field in the string
    Dim PosMessage As cBrokerMessage    ' Broker message object for the position
    
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        
        If (UCase(strFirstField) <> "BEGIN") And (UCase(strFirstField) <> "END") Then
            Set FillMessage = New cBrokerMessage
            Set PosMessage = New cBrokerMessage
            
            
            FillMessage.FromString strMessage
        
            PosMessage.Add "Account", FillMessage("Account")
            PosMessage.Add "Symbol", FillMessage("Symbol")
            
            m.Broker.PositionReceived PosMessage.ToString, False
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.GetPositionForFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConnectString
'' Description: Build a string to send for a Connect call
'' Inputs:      For Log?
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ConnectString(Optional ByVal bForLog As Boolean = False) As String
On Error GoTo ErrSection:

    Dim brokerMsg As cBrokerMessage     ' Broker Message object
    
    Set brokerMsg = New cBrokerMessage
    
    brokerMsg.Add "User", m.strUserName
    brokerMsg.Add "Pass", m.strPassword
    brokerMsg.Add "IP", m.strIP
    brokerMsg.Add "Port", m.strPort
    brokerMsg.Add "Target", m.strTargetID
    brokerMsg.Add "Subsender", m.strSubsender
    
    ConnectString = brokerMsg.ToString(bForLog)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.ConnectString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BrokerOrderString
'' Description: Build an order string to send to the broker
'' Inputs:      Order, New Order?, Previous Genesis ID
'' Returns:     Broker Order String
''
'' Fields:      GenesisID, BrokerID, Account, Symbol, Exchange, Order Type,
''              Buy/Sell, Quantity, Stop, Limit, TIF, Expiration
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BrokerOrderString(ByVal Order As cPtOrder, ByVal bNewOrder As Boolean, Optional ByVal strPreviousGenesisID As String = "") As String
On Error GoTo ErrSection:

    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim strBrokerSymbol As String       ' Broker symbol
    Dim strInstruction As String        ' Order Instruction to send
    
    strInstruction = ""
    If Order.OrderType = eTT_OrderType_MIT Then
        strInstruction = "MIT"
    End If
    
    strBrokerSymbol = BrokerSymbol(Order.Symbol, BrokerSym)
    BrokerOrderString = m.BrokerKv.OrderToMessage(Order, bNewOrder, strBrokerSymbol, BrokerSym, m.Broker, strPreviousGenesisID, strInstruction)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.BrokerOrderString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetLoginInfo
'' Description: Get login information from the user
'' Inputs:      User Name, Switching, Show IP?
'' Returns:     True if user OKed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetLoginInfo(Optional ByVal strUserName As String = "", Optional ByVal bSwitching As Boolean = False, Optional ByVal bShowIP As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = frmLoginFix.ShowMe(m.Broker, strUserName, bSwitching, bShowIP, False, True, True, , , m.astrEnabledSymbols)
    If bReturn = True Then
        m.strUserName = frmLoginFix.UserName
        m.strPassword = frmLoginFix.Password
        m.strIP = frmLoginFix.IP
        m.strPort = frmLoginFix.Port
        m.strTargetID = frmLoginFix.TargetID
        m.strSubsender = frmLoginFix.SubsenderID
        
        m.Broker.AskedPassword = True
    End If
    
    GetLoginInfo = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.GetLoginInfo"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleSymbols
'' Description: Handle a list of symbols enabled for market data from CQG
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleSymbols(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the string
    Dim brokerMessage As cBrokerMessage ' Broker message
    Dim strGenesisBase As String        ' Genesis base symbol
    Dim lPos As Long                    ' Position in the array
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrMessage As cGdArray         ' Message broken out into fields
    Dim strFileName As String           ' Filename for the LKSyms file
    
    If Len(strMessage) > 0 Then
        strFirstField = Parse(strMessage, vbTab, 1)
        If UCase(strFirstField) = "BEGIN" Then
            m.astrEnabledSymbols.Clear
        ElseIf UCase(strFirstField) = "END" Then
            m.Broker.DumpDebug vbTab & "Enabled Symbols = '" & m.astrEnabledSymbols.JoinFields(",") & "'"
            
            If m.astrEnabledSymbols.BinarySearch("AP") Then
                g.Broker.SfeAllowed(m.Broker.Broker) = Int(CurrentTime)
            Else
                g.Broker.SfeAllowed(m.Broker.Broker) = kNullData
            End If
            
            strFileName = LksFile
            
            m.astrEnabledSymbols.Serialize strFileName, True
            LKSCRC = CalcFileCrc(strFileName)
        Else
            Set astrMessage = New cGdArray
            astrMessage.SplitFields strMessage, vbTab
            
            For lIndex = 0 To astrMessage.Size - 1
                Set brokerMessage = New cBrokerMessage
                brokerMessage.FromString astrMessage(lIndex)
                
                strGenesisBase = GenesisBaseSymbol(brokerMessage("Symbol"))
                If Len(strGenesisBase) > 0 Then
                    If m.astrEnabledSymbols.BinarySearch(strGenesisBase, lPos) = False Then
                        m.astrEnabledSymbols.Add strGenesisBase, lPos
                    End If
                End If
            Next lIndex
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.HandleSymbols"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleRejectMessage
'' Description: Handle a reject message coming from CQG
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleRejectMessage(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim brokerMessage As cBrokerMessage ' Broker message
    
    If Len(strMessage) > 0 Then
        Set brokerMessage = New cBrokerMessage
        brokerMessage.FromString strMessage
        
        frmOnlineBroker.AddDialogMessage "Error received from " & m.Broker.BrokerName & ":||" & brokerMessage("Text") & "|", "i", , m.Broker.BrokerName & " Error"
        
        Select Case UCase(brokerMessage("MsgType"))
            Case "NEWORDERSINGLE", "NEWORDERLIST", "ORDERCANCELREQUEST", "ORDERCANCELREPLACEREQUEST"
                Refresh True
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.HandleRejectMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisBaseSymbol
'' Description: Determine the Genesis base symbol for the given CQG base symbol
'' Inputs:      Cqg Base Symbol, For Position?, Leg Gap (out), Broker Symbol (out)
'' Returns:     Genesis Base Symbol ( Blank if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GenesisBaseSymbol(ByVal strCqgBaseSymbol As String, Optional ByVal bForPosition As Boolean = False, Optional strLegGap As String = "", Optional BrokerSym As cBrokerSymbol) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrSymbol As cGdArray          ' CQG Symbol broken out into pieces
    Dim lIndex As Long                  ' Index into a for loop
    Dim strCqgBase As String            ' CQG base symbol from the symbols table
    Dim lLenCqgBase As Long             ' Length of the CQG base symbol
    Dim CqgSymbol As cBrokerSymbol      ' Broker symbol object
    Dim strBase As String               ' Base symbol
    
    strReturn = ""
    Set BrokerSym = Nothing
    
    If Len(strCqgBaseSymbol) > 0 Then
        Set astrSymbol = New cGdArray
        astrSymbol.SplitFields strCqgBaseSymbol, "."
        
        If astrSymbol.Size = 3 Then
            strBase = astrSymbol(2)
            
            If astrSymbol(0) = "F" Then
                For lIndex = 1 To m.Symbols.Count
                    Set CqgSymbol = m.Symbols(lIndex)
                    
                    If bForPosition Then
                        strCqgBase = CqgSymbol.BrokerPositionBase
                    Else
                        strCqgBase = CqgSymbol.BrokerBase
                    End If
                    lLenCqgBase = Len(strCqgBase)
                    
                    If InStr(strCqgBase, "S:") = 0 Then
                        If strCqgBase = strBase Then
                            Set BrokerSym = CqgSymbol
                            strReturn = BrokerSym.GenesisBase
                            
                            Exit For
                        End If
                    Else
                        strCqgBase = Replace(strCqgBase, "S:", "")
                        lLenCqgBase = Len(strCqgBase)
                    
                        If Len(strBase) > lLenCqgBase Then
                            If Left(strBase, lLenCqgBase) = strCqgBase Then
                                strLegGap = Right(strBase, Len(strBase) - lLenCqgBase)
                                If IsAlpha(strLegGap) = False Then
                                    Set BrokerSym = CqgSymbol
                                    strReturn = BrokerSym.GenesisBase
                                    
                                    Exit For
                                Else
                                    strLegGap = ""
                                End If
                            End If
                        End If
                    End If
                Next lIndex
            ElseIf astrSymbol(0) = "C" Or astrSymbol(0) = "P" Then
                strBase = "O:" & strBase
                
                If bForPosition Then
                    Set BrokerSym = m.Symbols.GetByBrokerPositionSymbol(strBase)
                Else
                    Set BrokerSym = m.Symbols.GetByBrokerSymbol(strBase)
                End If
                
                If Not BrokerSym Is Nothing Then
                    strReturn = BrokerSym.GenesisBase
                End If
            End If
        End If
    End If
    
    GenesisBaseSymbol = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerCqg.GenesisBaseSymbol"
    
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
    Set m.BrokerKv = New cBrokerKeyValue
    Set m.AmendOrders = New cGdTree
    Set m.Accounts = New cGdTree
    Set m.Orders = New cGdTree
    Set m.Fills = New cGdTree
    
    Set m.astrFillSymbols = New cGdArray
    m.astrFillSymbols.Create eGDARRAY_Strings

    Set m.astrEnabledSymbols = New cGdArray
    m.astrEnabledSymbols.Create eGDARRAY_Strings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.Class_Initialize"
    
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
    Set m.BrokerKv = Nothing
    Set m.AmendOrders = Nothing
    Set m.Accounts = Nothing
    Set m.Orders = Nothing
    Set m.Fills = Nothing
    Set m.astrFillSymbols = Nothing
    Set m.astrEnabledSymbols = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerCqg.Class_Terminate"
    
End Sub
