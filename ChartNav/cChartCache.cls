VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cChartCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    Chart As cChart
    
    oPatternProfit As cPatternProfit
    eOrdBarMode As eOrderBarMode
    eDetachStatus As enumDetachStatus
    nWindowState As Long
    
    nSymbolLinkColor As Long
    nPeriodLinkColor As Long
        
    strDetachedPlacment As String
    strNormalPlacement As String
    strRatioPlacement As String
End Type
Private m As mPrivate

Public Function CacheChartGet() As cChart
    Set CacheChartGet = m.Chart
End Function

Public Property Get CacheChartWindowState() As Long
    CacheChartWindowState = m.nWindowState
End Property

Public Sub CacheChartSave(ChartToCache As cChart)
On Error GoTo ErrSection:

    Dim i&
    Dim bFound As Boolean

    If ChartToCache Is Nothing Then Exit Sub
    If ChartToCache.Form Is Nothing Then Exit Sub

    Set m.Chart = ChartToCache
    
    'these are precautionary debugging statements (theoretically should never happen)
    If m.Chart.Tree Is Nothing Then
        DebugLog "Entry (cChartCache.CacheChartSave): Tree object is nothing - " & m.Chart.Symbol & ", " & m.Chart.Template & " (" & g.strChartPage & ")"
    Else
        For i = 1 To m.Chart.Tree.Count
            If TypeOf m.Chart.Tree(i) Is cIndicator Then
                If Not m.Chart.Tree(i).AlertObject Is Nothing Then
                    m.Chart.Tree(i).UpdateAlert 3, False        '6451
                    bFound = True
                End If
            End If
        Next
    End If
    
    If m.Chart.Annots Is Nothing Then
        DebugLog "Entry (cChartCache.CacheChartSave): Annots object is nothing - " & m.Chart.Symbol & ", " & m.Chart.Template & " (" & g.strChartPage & ")"
    Else
        For i = m.Chart.Annots.Count To 1 Step -1
            If Not m.Chart.Annots(i).AlertObject Is Nothing Then
                If m.Chart.Annots(i).AlertObject.AlertType = eGDAlertType_Price Then
                    m.Chart.Annots.Remove (i)           '6696
                Else
                    m.Chart.Annots(i).UpdateAlert 3, False        '6451
                End If
                bFound = True
            End If
        Next
    End If
    
    'save properties of form to chart object
    With ChartToCache.Form
        'disconnect chart object from form object
        .ClearChartObject
        
        m.nSymbolLinkColor = .WindowLink.SymbolColor
        m.nPeriodLinkColor = .WindowLink.PeriodColor
        
        m.strDetachedPlacment = .GetDetachedPlacement
        m.strNormalPlacement = .GetNormalPlacement
        m.strRatioPlacement = .GetRatioPlacement
        
        m.eDetachStatus = .DetachStatus
        m.eOrdBarMode = .OrderBarMode
        
        If m.eOrdBarMode = eOrdBarMode_PFP Then
            Set m.oPatternProfit = .PatternProfitObj
            'the grids are on the form, on restore forms are reused and this chart object
            'may not be restored to this same form so clear out the grids to free memory
            .PfpReset ePfpReset_ClearAll        'aardvark 6755
            .PatternProfitObj = Nothing
        End If
    
        m.nWindowState = .WindowState
    End With
    
    'disconnect form object from chart object
    Set m.Chart.Form = Nothing
    
    'clear last created & last edited annot
    m.Chart.LastEditCreate Nothing, False
    
    'set the flag to reload trades  -aardvark 7040
    m.Chart.SetTrackerTradesReload
    
    'these are precautionary debugging statements (theoretically should never happen)
    If m.Chart.Tree Is Nothing Then
        DebugLog "Exit (cChartCache.CacheChartSave): Tree object is nothing - " & m.Chart.Symbol & ", " & m.Chart.Template & " (" & g.strChartPage & ")"
    End If
    If m.Chart.Annots Is Nothing Then
        DebugLog "Exit (cChartCache.CacheChartSave): Annots object is nothing - " & m.Chart.Symbol & ", " & m.Chart.Template & " (" & g.strChartPage & ")"
    End If
    
    m.Chart.RemoveAllTickBuffers
    
    If bFound Then
        If FormIsLoaded("frmAlertsSetup") Then
            If frmAlertsSetup.Visible Then frmAlertsSetup.LoadGrid
        End If
    End If

'    m.Chart.geReleaseChart

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cChartCache.CacheChartSave"
    
End Sub

Public Function CacheChartRestore(frm As frmChart) As Boolean
On Error GoTo ErrSection:

    Dim i&
    Dim bFound As Boolean
    
    Dim Alert As cAlert

    Set m.Chart.Form = frm
    
    'precautionary debugging statements(theoretically should never happen)
    If m.Chart.Tree Is Nothing Then
        DebugLog "Chart object has no tree on restore."
    Else
        For i = 1 To m.Chart.Tree.Count
            If TypeOf m.Chart.Tree(i) Is cIndicator Then
                If Not m.Chart.Tree(i).AlertObject Is Nothing Then
                    m.Chart.Tree(i).UpdateAlert 1, False        '6451
                    bFound = True
                End If
            End If
        Next
    End If
    
    If m.Chart.Annots Is Nothing Then
        DebugLog "Chart object has no annots on restore."
    Else
        For i = m.Chart.Annots.Count To 1 Step -1           '6451
            Set Alert = m.Chart.Annots(i).AlertObject
            If Not Alert Is Nothing Then
                If m.Chart.Annots(i).AlertObject.AlertType = eGDAlertType_Price Then
                    m.Chart.Annots.Remove (i)       'theoretically should never get here
                Else
                    g.Alerts.Add Alert, Alert.ChartAlertId          '6698
                End If
                bFound = True
            End If
            
            If m.Chart.Annots(i).eType = eANNOT_ElliotLabel Then        '6926
                g.bPageHasEWILabels = True
            End If
        
        Next
        Set Alert = Nothing
    End If
    
    For i = 1 To g.Alerts.Count
        Set Alert = g.Alerts(i)
        If Not Alert Is Nothing Then
            If Alert.AlertType = eGDAlertType_Price Then
                If Alert.ShowOnCharts Then
                    If Alert.Symbol = m.Chart.Symbol Then
                        m.Chart.PriceAlertAdd Alert             '6696
                    End If
                End If
            End If
        End If
    Next
    
    Set Alert = Nothing
    
    With frm
        .SetChartObject m.Chart
        .WindowLink.SymbolColor = m.nSymbolLinkColor
        .WindowLink.PeriodColor = m.nPeriodLinkColor
        
        .SetDetachedPlacement m.strDetachedPlacment
        .SetNormalPlacement m.strNormalPlacement
        .SetRatioPlacement m.strRatioPlacement
        
        .DetachStatus = m.eDetachStatus
        .OrderBarMode = m.eOrdBarMode
        
        If .OrderBarMode = eOrdBarMode_PFP Then
            'init the grids         -aardvark 6755
            InitChartFlex frm, eFlexGridIdx_PfpInd
        End If
        
        'let the chart's timer miminize the form after all data members/objects are fully set/initialized
        'this is fix for grey & out-of-sync charts menu drop down reported by Vanessa
        If m.nWindowState = vbMinimized Then
            .tmr.Tag = "MINIMIZE_NOW"
        Else
            .WindowState = m.nWindowState
        End If
        
        .tmr.Enabled = False
    End With
    
    If m.Chart.IsProfileChart Then
        m.Chart.RedoMode = eRedo9_ReloadData    '6996
    Else
        m.Chart.TemplateLoadGlobalAnnots        '6447
        m.Chart.AddAllTickBuffers
        m.Chart.RedoMode = eRedo7_ReloadRT
    End If
    
    'precautionary debugging statements(theoretically should never happen)
    If m.Chart.Tree Is Nothing Or m.Chart.Annots Is Nothing Then
        DebugLog "Restore lost object on exit"
    End If

    If bFound Then
        If FormIsLoaded("frmAlertsSetup") Then
            If frmAlertsSetup.Visible Then frmAlertsSetup.LoadGrid
        End If
    End If
    
    'do this last so chart will not start processing PFP too soon
    If frm.OrderBarMode = eOrdBarMode_PFP Then
        frm.PatternProfitObj = m.oPatternProfit
    End If
    
    frm.SetAutoExit         '7040

ErrExit:
    CacheChartRestore = True
    Exit Function

ErrSection:
    RaiseError "cChartCache.CacheChartRestore"

End Function

Private Sub Class_Terminate()
    
    On Error Resume Next
    
    If Not m.Chart Is Nothing Then
        If Not m.Chart.Annots Is Nothing Then m.Chart.Annots.Clear
        If Not m.Chart.Tree Is Nothing Then m.Chart.Tree.Clear
        Set m.Chart = Nothing
    End If

End Sub

Public Sub CacheChartReleaseObj()

    Set m.Chart = Nothing       'this release reference to object; does not destroy it
End Sub
