VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAlerts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cAlerts.cls
'' Description: Collection of alerts
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/07/2009   DAJ         Changed delimiter for order info from | to ;
'' 01/12/2010   DAJ         Added num bars required for TradeSense alerts
'' 09/16/2010   DAJ         More information in e-mail order alerts (#5891)
'' 03/07/2011   DAJ         Broker Disconnect Alerts
'' 02/14/2012   DAJ         New status alerts for position mismatch / auto trade disabled
'' 08/03/2012   DAJ         Remove Photon
'' 08/07/2012   DAJ         Add PNL to alert with a fill
'' 10/16/2013   DAJ         Removed Xpress/Photon
'' 10/24/2014   DAJ         Fill Display
'' 12/02/2015   DAJ         Added Account filter for "Any Orders" alert
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kVersion = 4

Private Type mPrivate
    Alerts As cGdTree                   ' Collection of alerts
    strFileName As String               ' Name of the file to read the alerts from
    BarsColl As cGdTree                 ' Collection of Bars for streaming
    astrAlertMsgs As cGdArray           ' array of alert messages not yet saved to file
End Type

Private m As mPrivate

Public Property Let AlertMsgArray(aMsgs As cGdArray)
    Set m.astrAlertMsgs = aMsgs
End Property

Public Property Get AlertMsgArray() As cGdArray
    
    'This array is intended to be shared with frmAlertMessages. To ensure that modification to the array
    'is made only by frmAlertMessages when that form is loaded, the array object in this class is set to
    'nothing by frmAlertMessages.LoadGrid routine. The array object is returned by frmAlertMessages.Unload.
    Set AlertMsgArray = m.astrAlertMsgs

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Alerts = New cGdTree
    Set m.BarsColl = New cGdTree
    
gdResetProfiles 500, 510
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members upon termination
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    Set m.Alerts = Nothing
    Set m.astrAlertMsgs = Nothing
    
    For lIndex = 1 To m.BarsColl.Count
        g.RealTime.RemoveTickBuffer m.BarsColl(lIndex)
    Next lIndex
    Set m.BarsColl = Nothing
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.Class_Terminate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the alerts collection from the QuoteList.ALR file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load(ByVal strFileName As String)
On Error GoTo ErrSection:

    Dim astrAlertFile As New cGdArray   ' Array to hold the alerts file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim Alert As New cAlert             ' Alert object to add to the collection
    Dim astrFields As New cGdArray      ' Array to hold the quote board fields
    Dim lVersion As Long                ' Version of the file
    Dim strKey As String                'key into alert collections tree
    Dim bDeactivatedAlert As Boolean    ' Did we deactivate any alerts
    
    astrFields.SplitFields frmQuotes.AlertFields2, "|"

    m.strFileName = strFileName
    If astrAlertFile.FromFile(strFileName) = True Then
        If Parse(astrAlertFile(0), "=", 1) = "Version" Then
            lVersion = CLng(Val(Parse(astrAlertFile(0), "=", 2)))
        Else
            lVersion = 0&
        End If
        For lIndex = lVersion + 1 To kVersion
            ApplyVersion lIndex, astrAlertFile
        Next lIndex
        
        bDeactivatedAlert = False
        For lIndex = 1 To astrAlertFile.Size - 1
            Set Alert = New cAlert
            Alert.FromFileString astrAlertFile(lIndex)
            
            If UCase(Alert.Field) = "DATE" Then Alert.Field = "Session"
            
            ' Make sure to fill in the criteria id if applicable...
            If Len(Alert.CriteriaID) = 0 Then
                For lIndex2 = 0 To astrFields.Size - 1
                    If Parse(astrFields(lIndex2), ";", 1) = Alert.Field Then
                        Alert.CriteriaID = Parse(astrFields(lIndex2), ";", 3)
                        Exit For
                    End If
                Next lIndex2
            End If
            
            ' If there are TradeSense alerts that cannot have the number of required bars
            ' AND there is no override AND they are active, mark them as inactive
            ' and warn the user later
            If Alert.AlertType = eGDAlertType_TradeSense Then
                If (Alert.NumBarsCalc = -1&) And (Alert.NumBarsOver = -1&) And (Alert.Active = True) Then
                    Alert.Active = False
                    bDeactivatedAlert = True
                End If
            End If
            
            If IsValidAlert(Alert) = True Then
                strKey = Alert.AlertTypeText & Str(m.Alerts.Count + 1)
                m.Alerts.Add Alert, strKey
                Alert.AlertKey = strKey
            End If
        Next lIndex
        
        If bDeactivatedAlert = True Then
            InfBox "One or more TradeSense alerts have been deactivated because the number of required bars could not be calculated and there is no override", "!", , "TradeSense Alerts"
        End If
    End If
    
ErrExit:
    Set astrAlertFile = Nothing
    Set Alert = Nothing
    Exit Sub
    
ErrSection:
    Set astrAlertFile = Nothing
    Set Alert = Nothing
    RaiseError "cAlerts.Load", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the alerts collection to the QuoteList.ALR file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim fh As Integer                   ' File handle for an output file
    
    fh = FreeFile
    Open m.strFileName For Output As #fh
    Print #fh, "Version=" & Str(kVersion)
    For lIndex = 1 To m.Alerts.Count
        If Item(lIndex).AlertType <> eGDAlertType_Annot And Item(lIndex).AlertType <> eGDAlertType_Chart Then
            Print #fh, Item(lIndex).ToFileString
        End If
    Next lIndex
    Close #fh

    If Not m.astrAlertMsgs Is Nothing Then
        If m.astrAlertMsgs.Size > 0 Then
            
            Dim aFields As New cGdArray, aSave As New cGdArray
            
            'don't save alert key
            'dateTime \t alertType \t alertText \t alertKey \t alert time zone \t color
            For lIndex = 0 To m.astrAlertMsgs.Size - 1
                aFields.SplitFields m.astrAlertMsgs(lIndex), vbTab
                'dateTime \t alertType \t alertText \t alert time zone \t color
                aSave.Add aFields(0) & vbTab & aFields(1) & vbTab & aFields(2) & vbTab & aFields(4) & vbTab & aFields(5)
            Next
            
            If FileExist(AddSlash(App.Path) & kstrHistoryFile) Then
                aSave.ToFile AddSlash(App.Path) & kstrHistoryFile, True
            Else
                fh = FreeFile
                Open AddSlash(App.Path) & kstrHistoryFile For Append As #fh
                Print #fh, m.astrAlertMsgs(m.astrAlertMsgs.Size - 1)
                Close #fh
                aSave.ToFile AddSlash(App.Path) & kstrHistoryFile
            End If
            
            Set aFields = Nothing
            Set aSave = Nothing
            
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.Save", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Property:    Item (Get)
'' Description: Return the item at the given key or index
'' Inputs:      Key or Index of item to return
'' Returns:     Requested item if found, Nothing otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Item(ByVal vKey As Variant) As cAlert
Attribute Item.VB_UserMemId = 0
On Error GoTo ErrSection:

    If m.Alerts.Exists(vKey) Then
        Set Item = m.Alerts.Item(vKey)
    Else
        Set Item = Nothing
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAlerts.Item.Get", eGDRaiseError_Raise
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Property:    Item (Let)
'' Description: Set the item at the given Key or Index to the given Alert
'' Inputs:      Key or Index of item to replace, Alert to replace with
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Item(ByVal vKey As Variant, ByVal Alert As cAlert)
On Error GoTo ErrSection:

    Set m.Alerts.Item(vKey) = Alert

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAlerts.Item.Let", eGDRaiseError_Raise
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add the alert to the alerts collection
'' Inputs:      Alert to add
'' Returns:     Index where alert was added
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(ByVal Alert As cAlert, Optional strKey As String = "") As Long
On Error GoTo ErrSection:

    Dim i&

    If strKey = "" Then strKey = Alert.AlertTypeText & Str(m.Alerts.Count + 1)
    
    'if an alert was removed then the count would have changed and key of count+1 may have been assigned
    If Not m.Alerts(strKey) Is Nothing Then
        For i = 2 To 150
            strKey = Alert.AlertTypeText + Str(m.Alerts.Count + i)
            If m.Alerts(strKey) Is Nothing Then Exit For
        Next
    End If
    If m.Alerts(strKey) Is Nothing Then
        Alert.AlertKey = strKey
        Add = m.Alerts.Add(Alert, strKey)
        'If (Alert.AlertType = eGDAlertType_QuoteBoard) Then frmQuotes.DisplayAlert Alert, False
        
        'JM 02-09-2010: always call DisplayAlert so detached QB tabs & box style QB can be updated
        frmQuotes.DisplayAlert Alert, False, , , , True
        
        If Alert.AlertType = eGDAlertType_Price Then Alert.AddChartObject
    Else
        'we are way out of sync should restart
        InfBox "Alerts are out of sync. Please restart " & Parse(frmMain.Caption, "-", 1) & "."
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlerts.Add", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Remove
'' Description: Remove the element with the given key or index
'' Inputs:      Key or Index to remove
'' Returns:     True if item was found to remove, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Remove(ByVal vKey As Variant) As Boolean
On Error GoTo ErrSection:

    If m.Alerts.Exists(vKey) = True Then
        frmQuotes.DisplayAlert m.Alerts(vKey), True
        If FormIsLoaded("frmAlertMessages") Then frmAlertMessages.AlertRemoved m.Alerts(vKey)
    End If
    
    Remove = m.Alerts.Remove(vKey)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlerts.Remove", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckAlerts
'' Description: Check all of the active alerts in the collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckAlerts(Optional ByVal bShowStatus As Boolean = True, _
        Optional ByVal bOnlyBackColor As Boolean = False, _
        Optional ByVal bDisplayAlerts As Boolean = True, _
        Optional ByVal bIncludeTabAlerts As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Alert As cAlert
    
    'Dim strTabsDone As String
    'Dim bClearBackColor As Boolean
    
    Static bInProgress As Boolean       ' Are we currently in progress?
    
    If bInProgress = False Then
        bInProgress = True
        
        For lIndex = 1 To m.Alerts.Count
            Set Alert = Item(lIndex)
            If Alert.AlertType <> eGDAlertType_Annot And Alert.AlertType <> eGDAlertType_Chart Then
                If Alert.CheckFromCheckAlerts Then
                    If Alert.Active Or Alert.HasColorAction Then

                        'JM 02-16-2010: check tab alerts only from TotalRefresh since some QB (eg Bill Lopinto's) can take multiple seconds
                        If Len(Alert.TabName) = 0 Or bIncludeTabAlerts Or Not g.RealTime.Active Then
'JM 08-11-2010: cannot remember why this code is here, but it is the cause of issue 5861.
'                            If InStr(strTabsDone, Alert.TabName) Then
'                                bClearBackColor = False
'                            Else
'                                bClearBackColor = True
'                                strTabsDone = strTabsDone & "," & Alert.TabName
'                            End If

                            Alert.CheckAlert bOnlyBackColor         ', , , , bClearBackColor
                        End If
                    End If
                End If
            End If
        Next lIndex
        
        Set Alert = Nothing
        
        bInProgress = False
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.CheckAlerts", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckBrokerStatusAlert
'' Description: Check the given broker status alert in the collection
'' Inputs:      Broker
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CheckBrokerStatusAlert(ByVal nBroker As eTT_AccountType)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Alert As cAlert                 ' Alert object
    Static bInProgress As Boolean       ' Are we currently in progress?
    
    If bInProgress = False Then
        bInProgress = True
        
        For lIndex = 1 To m.Alerts.Count
            Set Alert = Item(lIndex)
            If Alert.Active = True Then
                If (Alert.AlertType = eGDAlertType_Status) And (Alert.Broker = nBroker) Then
                    Alert.CheckAlert
                End If
            End If
        Next lIndex
        
        bInProgress = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    bInProgress = False
    RaiseError "cAlerts.CheckBrokerStatusAlert"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Property:    Count (Get)
'' Description: Return the number of elements in the collection
'' Inputs:      None
'' Returns:     Number of elements in the collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Count() As Long
On Error GoTo ErrSection:

    If Not m.Alerts Is Nothing Then Count = m.Alerts.Count

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cAlerts.Count", eGDRaiseError_Raise
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveAlertsForSymbol
'' Description: Remove all alerts in the collection for the given symbol/period
'' Inputs:      Symbol/Period to remove alerts for
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveAlertsForSymbol(ByVal strSymbol As String, ByVal strPeriod As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bRemoved As Boolean             ' Did we remove anything?

    bRemoved = False
    For lIndex = m.Alerts.Count To 1 Step -1
        If Item(lIndex).Symbol = strSymbol And Item(lIndex).Period = strPeriod Then
            Remove lIndex
            bRemoved = True
        End If
    Next lIndex

    If bRemoved And FormIsLoaded("frmAlertsSetup") Then frmAlertsSetup.LoadGrid

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAlerts.RemoveAlertsForSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveAlertsForField
'' Description: Remove all alerts in the collection for the given field
'' Inputs:      Field to remove alerts for
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveAlertsForField(ByVal strField As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bRemoved As Boolean             ' Did we remove anything?

    bRemoved = False
    For lIndex = m.Alerts.Count To 1 Step -1
        If Item(lIndex).Field = strField Then
            Remove lIndex
            bRemoved = True
        End If
    Next lIndex

    If bRemoved And FormIsLoaded("frmAlertsSetup") Then frmAlertsSetup.LoadGrid

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cAlerts.RemoveAlertsForSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion
'' Description: Apply the updates for the given version
'' Inputs:      Version Number to Apply, Alert file
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion(ByVal lVersion As Long, astrAlertFile As cGdArray)
On Error GoTo ErrSection:

    Select Case lVersion
        Case 1
            ApplyVersion1 astrAlertFile
            
        Case 2
            ApplyVersion2 astrAlertFile
            
        Case 3
            ApplyVersion3 astrAlertFile
            
        Case 4
            ApplyVersion4 astrAlertFile
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.ApplyVersion"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion1
'' Description: Apply the updates to go from version 0 to version 1
'' Inputs:      Array of alerts
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion1(astrAlertFile As cGdArray)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrAlert As New cGdArray       ' Temporary array of alert information
    Dim strSymbol As String             ' Symbol
    Dim strPeriod As String             ' Period
    Dim lSymbolID As Long               ' Symbol ID for the symbol
    
    For lIndex = 0 To astrAlertFile.Size - 1
        astrAlert.SplitFields astrAlertFile(lIndex), "|"
        
        ' 1) Add the Alert Type as the new first element...
        astrAlert.Add Str(eGDAlertType_QuoteBoard), 0
        
        ' 2) Convert symbols to symbol ids if applicable...
        strSymbol = Parse(astrAlert(1), ";", 1)
        strPeriod = Parse(astrAlert(1), ";", 2)
        lSymbolID = GetSymbolID(strSymbol)
        If lSymbolID <> 0 Then astrAlert(1) = Str(lSymbolID) & ";" & strPeriod
        
        astrAlertFile(lIndex) = astrAlert.JoinFields("|")
    Next lIndex
    
    ' 3) Add a version line to the beginning...
    astrAlertFile.Add "Version=1", 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.ApplyVersion1", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion2
'' Description: Apply the updates to go from version 1 to version 2
'' Inputs:      Array of alerts
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion2(astrAlertFile As cGdArray)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim astrAlert As New cGdArray       ' Temporary array of alert information
    Dim astrNewAlert As New cGdArray    ' Temporary array of alert information
    Dim nAlertType As eGDAlertType      ' Alert type for the alert
    Dim astrActions As New cGdArray     ' Array of actions for an alert
    Dim nActionType As eGDAlertAction   ' Action for the alert
    
    Dim strSymbol$, strPeriodicity$
    Dim iSymbolID&
    
    astrActions.Create eGDARRAY_Strings, eAA_NumActions
    
    For lIndex = astrAlertFile.Size - 1 To 1 Step -1
        astrNewAlert.Clear
        astrActions(eAA_MessageBox) = "0,"
        astrActions(eAA_LogToFile) = "0,,"
        astrActions(eAA_ChangeBackColor) = "0,"
        astrActions(eAA_SendPage) = "0,"
        astrActions(eAA_SendEmail) = "0,,,,"
        astrActions(eAA_PlaceOrder) = "0,"
        astrActions(eAA_PlaySound) = "0,"
        
        astrAlert.SplitFields astrAlertFile(lIndex), "|"
        nAlertType = CLng(Val(astrAlert(0)))
        Select Case nAlertType
            Case eGDAlertType_QuoteBoard
                strSymbol = Parse(astrAlert(1), ";", 1)
                strPeriodicity = Parse(astrAlert(1), ";", 2)
            
                If Len(strSymbol) > 0 And Len(strPeriodicity) > 0 Then
                    If IsDigit(strSymbol) Then
                        iSymbolID = ValOfText(strSymbol)
                        If Len(g.SymbolPool.SymbolForID(iSymbolID)) = 0 Then
                            astrAlertFile.Remove lIndex     '40942 does not yield a valid symbol (HU)
                            iSymbolID = 0
                        End If
                    Else
                        iSymbolID = g.SymbolPool.SymbolIDforSymbol(strSymbol)
                        If iSymbolID > 0 Then
                            astrAlert(1) = iSymbolID & ";" & strPeriodicity
                        Else
                            astrAlertFile.Remove lIndex     'ZE-057 does not yield a valid symbolID
                        End If
                    End If
                
                    If iSymbolID > 0 Then
                        nActionType = CLng(Val(astrAlert(5)))
                        Select Case nActionType
                            Case eAA_MessageBox
                                astrActions(eAA_MessageBox) = "1," & astrAlert(6)
                                If Len(astrAlert(12)) > 0 Then
                                    astrActions(eAA_PlaySound) = "1," & astrAlert(12)
                                End If
                                
                            Case eAA_LogToFile
                                astrActions(eAA_LogToFile) = "1," & astrAlert(6) & "," & astrAlert(7)
                            
                            Case eAA_ChangeBackColor
                                astrActions(eAA_ChangeBackColor) = "1," & astrAlert(9)
                            
                            Case eAA_SendPage
                                astrActions(eAA_SendPage) = "1," & astrAlert(6)
                            
                            Case eAA_SendEmail
                                astrActions(eAA_SendEmail) = "1," & astrAlert(6) & "," & astrAlert(13) & "," & astrAlert(14) & "," & astrAlert(15)
                            
                            Case eAA_PlaceOrder
                                astrActions(eAA_PlaceOrder) = "1," & astrAlert(16)
                            
                            Case eAA_PlaySound
                                astrActions(eAA_PlaySound) = "1," & astrAlert(12)
                        
                        End Select
                
                        astrActions(eAA_MsgHistory) = "1"
                        
                        For lIndex2 = 0 To 4
                            astrNewAlert(lIndex2) = astrAlert(lIndex2)
                        Next lIndex2
                        astrNewAlert.Add "1", 1
                        astrNewAlert.Add "", 4
                        astrNewAlert(7) = astrAlert(10)
                        astrNewAlert(8) = astrAlert(8)
                        astrNewAlert(9) = astrAlert(11)
                        astrNewAlert(10) = astrAlert(17)
                        astrNewAlert(11) = astrActions.JoinFields("~")
                        
                        astrAlertFile(lIndex) = astrNewAlert.JoinFields("|")
                    End If
                End If
            
            Case eGDAlertType_AutoTrade
                nActionType = CLng(Val(astrAlert(3)))
                Select Case nActionType
                    Case eAA_MessageBox
                        astrActions(eAA_MessageBox) = "1," & astrAlert(4)
                        If Len(astrAlert(8)) > 0 Then
                            astrActions(eAA_PlaySound) = "1," & astrAlert(8)
                        End If
                        
                    Case eAA_LogToFile
                        astrActions(eAA_LogToFile) = "1," & astrAlert(4) & "," & astrAlert(5)
                    
                    Case eAA_SendPage
                        astrActions(eAA_SendPage) = "1," & astrAlert(4)
                    
                    Case eAA_SendEmail
                        astrActions(eAA_SendEmail) = "1," & astrAlert(4) & "," & astrAlert(9) & "," & astrAlert(10) & "," & astrAlert(11)
                    
                    Case eAA_PlaySound
                        astrActions(eAA_PlaySound) = "1," & astrAlert(8)
                
                End Select
                
                astrActions(eAA_MsgHistory) = "1"
                
                astrNewAlert(0) = astrAlert(0)
                astrNewAlert(1) = "-1"
                astrNewAlert(2) = astrAlert(1)
                astrNewAlert(3) = astrAlert(2)
                astrNewAlert(4) = astrAlert(7)
                astrNewAlert(5) = astrAlert(6)
                astrNewAlert(6) = astrActions.JoinFields("~")
            
                astrAlertFile(lIndex) = astrNewAlert.JoinFields("|")
                
            Case eGDAlertType_Status
                nActionType = CLng(Val(astrAlert(2)))
                Select Case nActionType
                    Case eAA_MessageBox
                        astrActions(eAA_MessageBox) = "1," & astrAlert(3)
                        If Len(astrAlert(7)) > 0 Then
                            astrActions(eAA_PlaySound) = "1," & astrAlert(7)
                        End If
                        
                    Case eAA_LogToFile
                        astrActions(eAA_LogToFile) = "1," & astrAlert(3) & "," & astrAlert(4)
                    
                    Case eAA_SendPage
                        astrActions(eAA_SendPage) = "1," & astrAlert(3)
                    
                    Case eAA_SendEmail
                        astrActions(eAA_SendEmail) = "1," & astrAlert(3) & "," & astrAlert(8) & "," & astrAlert(9) & "," & astrAlert(10)
                    
                    Case eAA_PlaySound
                        astrActions(eAA_PlaySound) = "1," & astrAlert(7)
                
                End Select
                
                astrActions(eAA_MsgHistory) = "1"
            
                astrNewAlert(0) = astrAlert(0)
                astrNewAlert(1) = astrAlert(1)
                astrNewAlert(2) = astrAlert(6)
                astrNewAlert(3) = astrAlert(5)
                astrNewAlert(4) = astrActions.JoinFields("~")
            
                astrAlertFile(lIndex) = astrNewAlert.JoinFields("|")
                
        End Select
    Next lIndex
    
    astrAlertFile(0) = "Version=2"
    astrAlertFile.ToFile m.strFileName

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.ApplyVersion2"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion3
'' Description: Apply the updates to go from version 2 to version 3
'' Inputs:      Array of alerts
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion3(astrAlertFile As cGdArray)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrAlert As cGdArray           ' Alert broken up into an array
    Dim nAlertType As eGDAlertType      ' Type of the alert
    Dim Alert As cAlert                 ' Alert object
    
    For lIndex = 0 To astrAlertFile.Size - 1
        Set astrAlert = New cGdArray
        astrAlert.SplitFields astrAlertFile(lIndex), "|"
        
        nAlertType = CLng(Val(astrAlert(0)))
        
        If nAlertType = eGDAlertType_TradeSense Then
            If astrAlert.Size = 9 Then
                Set Alert = New cAlert
                Alert.FromFileString astrAlertFile(lIndex)
                Alert.AutoDetectNumBars
                astrAlertFile(lIndex) = Alert.ToFileString
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.ApplyVersion3"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion4
'' Description: Apply the updates to go from version 3 to version 4
'' Inputs:      Array of alerts
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion4(astrAlertFile As cGdArray)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim astrAlert As cGdArray           ' Alert broken up into an array
    Dim nAlertType As eGDAlertType      ' Type of the alert
    Dim Alert As cAlert                 ' Alert object
    Dim strBrokerName As String         ' Broker Name
    Dim bFound As Boolean               ' Was the broker name found?
    
    For lIndex = astrAlertFile.Size - 1 To 0 Step -1
        Set astrAlert = New cGdArray
        astrAlert.SplitFields astrAlertFile(lIndex), "|"
        
        nAlertType = CLng(Val(astrAlert(0)))
        
        If nAlertType = eGDAlertType_Status Then
            If InStr(astrAlert(1), ";") = 0 Then
                If (InStr(astrAlert(1), "Online Brokerage") > 0) Then
                    bFound = False
                    strBrokerName = Replace(astrAlert(1), " Online Brokerage", "")
                    
                    Select Case UCase(strBrokerName)
'                        Case "LIND WALDOCK"
'                            astrAlert(1) = astrAlert(1) & ";" & Str(eTT_AccountType_LindWaldock)
'                            astrAlertFile(lIndex) = astrAlert.JoinFields("|")
                        
'                        Case "PHOTON"
'                            astrAlert(1) = astrAlert(1) & ";" & Str(3) ' eTT_AccountType_Photon
'                            astrAlertFile(lIndex) = astrAlert.JoinFields("|")
                        
                    End Select
                Else
                    astrAlert(1) = astrAlert(1) & ";-1"
                    astrAlertFile(lIndex) = astrAlert.JoinFields("|")
                End If
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.ApplyVersion4"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddAlertMessage
'' Description: Add a message to the current message array
'' Inputs:      Date, Message, Alert Object, Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddAlertMessage(ByVal dDate As Double, ByVal strMessage As String, Alert As cAlert, Optional ByVal strSymbol As String = "")
On Error GoTo ErrSection:

    If Alert Is Nothing Then Exit Sub
    
    If FormIsLoaded("frmAlertMessages") Then
        frmAlertMessages.AddAlertMessage dDate, strMessage, Alert, strSymbol
    Else
        If m.astrAlertMsgs Is Nothing Then Set m.astrAlertMsgs = New cGdArray

        m.astrAlertMsgs.Add Str(dDate) & vbTab & Alert.AlertTypeText & vbTab _
            & strMessage & vbTab & Alert.AlertKey & vbTab & Alert.LastCheckedTimeZone & vbTab & Alert.MessageColor & vbTab & strSymbol
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.AddAlertMessage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveAlertsForManualOrder
'' Description: Remove any alerts for the given order ID
'' Inputs:      Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveAlertsForManualOrder(ByVal lOrderID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bRemoved As Boolean             ' Did we remove anything?
    
    bRemoved = False
    For lIndex = m.Alerts.Count To 1 Step -1
        If Item(lIndex).AlertType = eGDAlertType_AutoTrade Then
            If ((Item(lIndex).OrderAlertType = eGDOrderAlertType_ManualOrder) And (Item(lIndex).TradeItemID = lOrderID)) Then
                m.Alerts.Remove lIndex
                bRemoved = True
            End If
        End If
    Next lIndex
    
    If bRemoved And FormIsLoaded("frmAlertsSetup") Then frmAlertsSetup.LoadGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.RemoveAlertsForManualOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveAlertsForAutoTradeItem
'' Description: Remove any alerts for the given automated trading item ID
'' Inputs:      Automated Trading Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RemoveAlertsForAutoTradeItem(ByVal lAutoTradeItemID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bRemoved As Boolean             ' Did we remove anything?
    
    For lIndex = m.Alerts.Count To 1 Step -1
        If Item(lIndex).AlertType = eGDAlertType_AutoTrade Then
            If ((Item(lIndex).OrderAlertType = eGDOrderAlertType_AutoTrade) And (Item(lIndex).TradeItemID = lAutoTradeItemID)) Then
                m.Alerts.Remove lIndex
                bRemoved = True
            End If
        End If
    Next lIndex

    If bRemoved And FormIsLoaded("frmAlertsSetup") Then frmAlertsSetup.LoadGrid

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.RemoveAlertsForAutoTradeItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsValidAlert
'' Description: Is the alert still valid? (e.g. does the order still exist?)
'' Inputs:      Alert
'' Returns:     True if Valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsValidAlert(ByVal Alert As cAlert) As Boolean
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Temporary order object
    Dim TradeItem As New cAutoTradeItem ' Temporary trading item object
    Dim bReturn As Boolean              ' Return value for the function

    bReturn = True
    With Alert
        Select Case .AlertType
            Case eGDAlertType_AutoTrade
                Select Case .OrderAlertType
                    Case eGDOrderAlertType_AutoTrade
                        If TradeItem.Load(.TradeItemID) = False Then
                            bReturn = False
                        End If
                    Case eGDOrderAlertType_ManualOrder
                        If Order.Load(.TradeItemID) = True Then
                            bReturn = IsOpenOrder(Order.Status)
                        Else
                            bReturn = False
                        End If
                    Case eGDOrderAlertType_AllOrders
                        bReturn = True
                End Select
            Case eGDAlertType_QuoteBoard
                If Alert.IsSymbol Then
                    If Len(Alert.Symbol) = 0 Or Alert.SymbolID = 0 Then
                        bReturn = False
                    End If
                End If
            Case eGDAlertType_Chart, eGDAlertType_Annot
                If Len(Alert.Symbol) = 0 Or Alert.SymbolID = 0 Then
                    bReturn = False
                End If
            Case eGDAlertType_Status
                If .Broker > 0 Then
                    bReturn = g.Broker.IsBrokerUser(.Broker)
                End If
        End Select
    End With
    
    IsValidAlert = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlerts.IsValidAlert"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderStatusChange
'' Description: An order status has changed so check the alerts
'' Inputs:      Order that Changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OrderStatusChange(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrOrderInfo As New cGdArray   ' Order information to pass to the alert
    Dim Fill As cPtFill                 ' Fill for the order
    Dim strPnl As String                ' Profit/Loss information for the fill
    
    If Len(Order.BrokerID) = 0 Then
        astrOrderInfo(0) = "Order # : " & Str(Order.OrderID)
    Else
        astrOrderInfo(0) = "Order # : " & Order.BrokerID
    End If
    astrOrderInfo(1) = "Order   : " & Order.OrderText(True, True)
    astrOrderInfo(2) = "Status  : " & OrderStatus(Order.Status)
    If Len(Order.Message) > 0 Then
        astrOrderInfo(3) = "Message : " & Order.Message
    End If
    If (Order.FillQuantity > 0) Then
        For lIndex = 1 To Order.Fills.Count
            Set Fill = Order.Fills(lIndex)
            astrOrderInfo.Add "Fill    : " & mTradeTracker.FillDisplay(Fill, False, True, False, False, False, False, True)
        Next lIndex
    End If
    
    For lIndex = 1 To m.Alerts.Count
        If (Item(lIndex).Active = True) Then
            If (Item(lIndex).IsOrderChangeStatusAlert = True) Then
                Item(lIndex).CheckAlert False, astrOrderInfo.JoinFields(";")
            ElseIf (Item(lIndex).AlertType = eGDAlertType_AutoTrade) And (Item(lIndex).OrderAlertType = eGDOrderAlertType_AllOrders) Then
                If Item(lIndex).ValidAccountForAutoTradeAlert(Order.AccountID, Order.Broker) Then
                    Item(lIndex).CheckAlert False, astrOrderInfo.JoinFields(";")
                End If
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.OrderStatusChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PositionMismatch
'' Description: A position mismatch has occurred
'' Inputs:      Symbol, Account
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PositionMismatch(ByVal strSymbol As String, ByVal strAccount As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim StatusInfo As cStatusAlertInfo  ' Status Alert information
    
    For lIndex = 1 To m.Alerts.Count
        If (Item(lIndex).Active = True) Then
            If (Item(lIndex).IsPositionMismatchStatusAlert = True) Then
                Set StatusInfo = New cStatusAlertInfo
                
                StatusInfo.Symbol = strSymbol
                StatusInfo.Account = strAccount
                
                Item(lIndex).CheckAlert False, , , , StatusInfo
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.PositionMismatch"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoTradeItemDisabled
'' Description: An automated trading item has been disabled
'' Inputs:      Symbol, Account, Auto Trade Item, Reason
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AutoTradeItemDisabled(ByVal strSymbol As String, ByVal strAccount As String, ByVal strAutoTradeItem As String, ByVal strReason As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim StatusInfo As cStatusAlertInfo  ' Status Alert information
    
    For lIndex = 1 To m.Alerts.Count
        If (Item(lIndex).Active = True) Then
            If (Item(lIndex).IsAutoTradeItemDisabledStatusAlert = True) Then
                Set StatusInfo = New cStatusAlertInfo
                
                StatusInfo.Symbol = strSymbol
                StatusInfo.Account = strAccount
                StatusInfo.AutoTradeItem = strAutoTradeItem
                StatusInfo.Reason = strReason
                
                Item(lIndex).CheckAlert False, , , , StatusInfo
            End If
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.AutoTradeItemDisabled"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadBars
'' Description: Load and splice the bars for a particular symbol or symbol ID
'' Inputs:      Bars to Load, Symbol or Symbol ID to load, Period, Num Bars
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadBars(Bars As cGdBars, ByVal vSymbolOrSymbolID As Variant, ByVal strPeriod As String, Optional ByVal lNumBars As Long = -1&) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value
    Dim lStartDate As Long              ' Start date value to pass to GetBars
    
    If lNumBars = -1& Then
        lStartDate = LastDailyDownload - 5
    Else
        lStartDate = lNumBars * -1&
    End If
    
    Bars.ArrayMask = eBARS_EodBidAsk
    bReturn = DM_GetBars(Bars, vSymbolOrSymbolID, strPeriod, lStartDate, , , False)
    
    g.RealTime.AddTickBuffer Bars, False ' TLB: don't need to add to recent symbols list (esp. if from quote board)
    g.RealTime.SpliceBars Bars
    
    LoadBars = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlerts.LoadBars"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddBars
'' Description: Add a symbol/period combination to the bars collection
'' Inputs:      Symbol or Symbol ID, Period, Num Bars
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddBars(ByVal vSymbolOrSymbolID As Variant, ByVal strPeriod As String, Optional ByVal lNumBars As Long = -1&)
On Error GoTo ErrSection:

    Dim Bars As New cGdBars             ' Temporary Bars structure

    If m.BarsColl.Exists(BarsKey(vSymbolOrSymbolID, strPeriod)) = False Then
        LoadBars Bars, vSymbolOrSymbolID, strPeriod, lNumBars
        m.BarsColl.Add Bars, BarsKey(vSymbolOrSymbolID, strPeriod)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.AddBars"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetBars
'' Description: Get the Bars for the given symbol from the collection (add it
''              to the collection if not there)
'' Inputs:      Symbol, Period, Num Bars
'' Returns:     Bars
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetBars(ByVal vSymbolOrSymbolID As Variant, ByVal strPeriod As String, Optional ByVal lNumBars As Long = -1&) As cGdBars
On Error GoTo ErrSection:

    If m.BarsColl.Exists(BarsKey(vSymbolOrSymbolID, strPeriod)) = False Then
        AddBars vSymbolOrSymbolID, strPeriod, lNumBars
    End If
    
    Set GetBars = m.BarsColl(BarsKey(vSymbolOrSymbolID, strPeriod))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlerts.GetBars"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateBars
'' Description: Update the Bars from the stream
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateBars()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bNewBar As Boolean              ' There is a new bar, so we need to reload data
    Dim bBarsReloaded As Boolean
    Dim Bars As cGdBars
    
    bBarsReloaded = False
    For lIndex = 1 To m.BarsColl.Count
        If g.RealTime.Active Then
            If g.RealTime.UpdateBars(m.BarsColl(lIndex), bNewBar) = True Then
                If bNewBar Then
                    ' For this purpose, all we need to do is reload today's data which is what
                    ' SpliceBars does, so we don't need to reload all of history (06/11/2008 DAJ)...
                    g.RealTime.SpliceBars m.BarsColl(lIndex)
                    
                    'If m.BarsColl(lIndex).Prop(eBARS_SymbolID) = 0 Then
                    '    LoadBars m.BarsColl(lIndex), m.BarsColl(lIndex).Prop(eBARS_Symbol), m.BarsColl(lIndex).Prop(eBARS_PeriodicityStr)
                    'Else
                    '    LoadBars m.BarsColl(lIndex), m.BarsColl(lIndex).Prop(eBARS_SymbolID), m.BarsColl(lIndex).Prop(eBARS_PeriodicityStr)
                    'End If
                    
                    bBarsReloaded = True
                End If
            End If
        Else
            Set Bars = m.BarsColl(lIndex)
            Bars.Size = 0
            LoadBars Bars, Bars.Prop(eBARS_SymbolID), Bars.Prop(eBARS_PeriodicityStr)       '5200
        End If
    Next lIndex
    
    If bBarsReloaded Then
        DebugLog gdGetProfiles(500, 500)
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.UpdateBars"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BarsKey
'' Description: Determine the key into the bars collection
'' Inputs:      Symbol or Symbol ID, Period
'' Returns:     Key into the collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BarsKey(ByVal vSymbolOrSymbolID As Variant, ByVal strPeriod As String) As String
On Error GoTo ErrSection:

    BarsKey = Str(vSymbolOrSymbolID) & vbTab & strPeriod

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlerts.BarsKey"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayQBAlerts
'' Description: Display all of the quote board alerts for the current tab
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisplayQBAlerts(Optional ByRef frm As frmDetachedQBTab)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRedraw As Long
        
    If frm Is Nothing Then
        If frmQuotes.CurrentTabStyle(True) = eGDQuoteStyle_Grid Then
            'Note: the redraw mode is set here for performance reason.
            'Bill Lopinto test case:
            '   Each time the frmQuotes.DisplayAlert is called it searches a table of 200 records to locate
            '   the cell in the grid matching alert object properties. Bill Lopinto has 528 alerts.
            '   When redraw mode is set by the frmQuotes.DisplayAlert profiling gives average of 0.60 seconds.
            '   Overiding redraw mode here results in profiling average of 0.09 second
            
            lRedraw = frmQuotes.fgQuotes.Redraw
            frmQuotes.fgQuotes.Redraw = flexRDNone
            For lIndex = 1 To Count
                If Item(lIndex).AlertType = eGDAlertType_QuoteBoard Then
                    frmQuotes.DisplayAlert Item(lIndex), , False, , True
                    If Not g.bStarting And Not g.bUnloading And Not g.RealTime.Active Then
                        If Len(Item(lIndex).TabName) > 0 Then Item(lIndex).QBTabAlertColorOnly frmQuotes
                    End If
                End If
            Next lIndex
            
            frmQuotes.fgQuotes.Redraw = lRedraw
        Else
            frmQuotes.InitBoxQbAlerts
        End If
    Else
        lRedraw = frm.fgQuotes.Redraw
        frm.fgQuotes.Redraw = flexRDNone
        For lIndex = 1 To Count
            If Item(lIndex).AlertType = eGDAlertType_QuoteBoard Then
                frm.DisplayAlert Item(lIndex), False
                If Not g.bStarting And Not g.bUnloading And Not g.RealTime.Active Then
                    If Len(Item(lIndex).TabName) > 0 Then Item(lIndex).QBTabAlertColorOnly frm
                End If
            End If
        Next lIndex

        frm.fgQuotes.Redraw = lRedraw
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAlerts.DisplayQBAlerts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AlertExists
'' Description: Does the given alert exist in the collection?
'' Inputs:      Alert
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AlertExists(ByVal strAlert As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    For lIndex = 1 To m.Alerts.Count
        If Item(lIndex).IsTheSame(strAlert) Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    AlertExists = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAlerts.AlertExists"
    
End Function
