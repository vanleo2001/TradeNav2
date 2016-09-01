VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRptBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit

Private Type mPrivate
    RptForm As Form                     ' Report form for this bridge
End Type
Private m As mPrivate

Property Get AppPath() As String
    AppPath = m.RptForm.AppPath
End Property
Property Let AppPath(ByVal strPath As String)
    m.RptForm.AppPath = strPath
    g.strAppPath = strPath
End Property

Property Get ImageList() As ListImages
    Set ImageList = m.RptForm.Images
End Property
Property Let ImageList(pData As ListImages)
    m.RptForm.Images = pData
End Property

Property Let MainForm(pData As Object)
    m.RptForm.MainForm = pData
End Property

Property Get Pyramiding() As Boolean
    Pyramiding = m.RptForm.Pyramiding
End Property
Property Let Pyramiding(ByVal pData As Boolean)
    m.RptForm.Pyramiding = pData
End Property

Property Get DLLForm() As Object
    Set DLLForm = m.RptForm
End Property
Property Let DB(pData As Database)
    m.RptForm.DB = pData
End Property
Property Let AppName(pData As String)
    m.RptForm.AppName = pData
End Property
Property Let PortOrSystemName(pData As String)
    m.RptForm.SystemName = pData
End Property
Property Let FromDate(pData As Double)
    m.RptForm.FromDate = pData
End Property
Property Let ToDate(pData As Double)
    m.RptForm.ToDate = pData
End Property
Property Let DefaultBeginBalance(pData As Double)
    m.RptForm.DefaultBeginBalance = pData
End Property
Property Let SetSystemsListHandle(pHandle As Long)
    m.RptForm.SystemsListHandle = pHandle
End Property
Public Sub SetTradeHandles(pTradesDtlHandle As Long, pTradesHdrHandle As Long)
    m.RptForm.TradesHandles pTradesDtlHandle, pTradesHdrHandle
End Sub
Public Property Get ChartHwnd() As Long
    ChartHwnd = m.RptForm.ChartHwnd
End Property
Public Property Let ChartHwnd(ByVal hChartHwnd As Long)
    m.RptForm.ChartHwnd = hChartHwnd
End Property
Public Property Get Help() As Object
    'Help = m.RptForm.Help
    Set Help = g.Help
End Property
Public Property Let Help(HelpObj As Object)
    'm.RptForm.Help = HelpObj
    Set g.Help = HelpObj
End Property
Public Property Get IsLoaded() As Boolean
    IsLoaded = m.RptForm.IsLoaded
End Property

Public Property Get ShowInLocalTime() As Boolean
    ShowInLocalTime = g.bShowInLocalTimeZone
End Property
Public Property Let ShowInLocalTime(ByVal bShowInLocalTime As Boolean)
    g.bShowInLocalTimeZone = bShowInLocalTime
End Property

Public Property Get AltGridRowColor() As Long
    AltGridRowColor = g.nAltGridRowColor
End Property
Public Property Let AltGridRowColor(ByVal nAltGridRowColor As Long)
    g.nAltGridRowColor = nAltGridRowColor
End Property

Public Property Get HideTdoReports() As Boolean
    HideTdoReports = m.RptForm.HideTdoReports
End Property
Public Property Let HideTdoReports(ByVal bHideTdoReports As Boolean)
    m.RptForm.HideTdoReports = bHideTdoReports
End Property

Public Sub SetAppBackColor(ByVal nBackColor As Long)
    On Error Resume Next
    mGenesis.SetAppBackColor nBackColor
End Sub

Public Sub Show(Optional bNonModal As Boolean = False)
On Error GoTo ErrSection:
    
    'Validations
    If m.RptForm.Trades Is Nothing Then
        InfBox "No trade history was found.  Reports cannot be viewed.", "!", , "Reports Error"
        Exit Sub
    End If
    If m.RptForm.Trades.NumRecords - 1 <= 0 Then
        InfBox "No trade history was found.  Reports cannot be viewed.", "!", , "Reports Error"
        Exit Sub
    End If
    If (m.RptForm.Trades.NumRecords = 3) And (m.RptForm.Trades.Item(2, entd_RuleID) = 0) Then
        InfBox "There was only one entry with no exit.  Reports cannot be viewed.", "!", , "Reports Error"
        Exit Sub
    End If
    
    If Not bNonModal Then
        ShowForm m.RptForm, True
    Else
        MoveFormOnScreen m.RptForm
        SetWindowPos m.RptForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptBridge.Show", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

Public Sub ImportTrades(ByVal strTradesPath$, ByVal lSystemNumber&, ByVal strSystemName$)
On Error GoTo ErrSection:

    m.RptForm.ImportTrades strTradesPath, lSystemNumber, strSystemName

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cRptBridge.ImportTrades", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

Public Sub ImportWithHandles(ByVal strTradesPath$, ByVal lSystemNumber&, _
                    ByVal strSystemName$, ByVal hTblRules&, ByVal hTblMarkets&)
On Error GoTo ErrSection:

    m.RptForm.ImportTrades strTradesPath, lSystemNumber, strSystemName, hTblRules, hTblMarkets

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptBridge.ImportWithHandles", eGDRaiseError_Raise, m.RptForm.AppPath

End Sub

Public Sub ImportMultipleWithHandles(ByVal hFiles As Long, ByVal hTblRules As Long)
On Error GoTo ErrSection:

    m.RptForm.ImportMultipleTradeFiles hFiles, hTblRules

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRptBridge.ImportMultipleWithHandles", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CaptureReport
'' Description: Capture a report and save it to the given file
'' Inputs:      Report, Filename
'' Returns:     True if success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CaptureReport(ByVal lReport As Long, ByVal strFileName As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If m.RptForm.Trades Is Nothing Then
        InfBox "No trade history was found.  Reports cannot be viewed.", "!", , "Reports Error"
    ElseIf m.RptForm.Trades.NumRecords - 1 <= 0 Then
        InfBox "No trade history was found.  Reports cannot be viewed.", "!", , "Reports Error"
    ElseIf (m.RptForm.Trades.NumRecords = 3) And (m.RptForm.Trades.Item(2, entd_RuleID) = 0) Then
        InfBox "There was only one entry with no exit.  Reports cannot be viewed.", "!", , "Reports Error"
    Else
        Load m.RptForm
        
        m.RptForm.vsReport.SaveGrid strFileName, flexFileAll
        
        bReturn = True
    End If
    
    CaptureReport = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRptBridge.CaptureReport", , g.strAppPath

End Function

Private Sub Class_Initialize()

    Set m.RptForm = New frmReports
    
    If g.nAltGridRowColor = 0 Then
        g.nAltGridRowColor = &HE0FCFF
    End If

End Sub

Private Sub Class_Terminate()
    
    If Not m.RptForm Is Nothing Then Unload m.RptForm
    Set m.RptForm = Nothing
    
End Sub

Public Sub SetIrxBars(ByVal hIrxBars As Long)
    Set g.IrxBars = New cGdBars
    g.IrxBars.CopyFromHandle hIrxBars
End Sub

