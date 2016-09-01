VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    StatusBar As Object
    lSteps As Long
    lStep As Long
    dLastGridRefresh As Double
    bUseMessage As Boolean
End Type
Private m As mPrivate

Property Let StatusBarControl(pData As Object)
On Error GoTo ErrSection:

    Set m.StatusBar = pData
    With m.StatusBar
        .Appearance = flex3DLight
        .Cell(flexcpFontBold, 0, 0) = True
        .Cell(flexcpForeColor, 0, 0) = &HFFFFFF   'vbWhite
        .Cell(flexcpBackColor, 0, 0) = &H808080
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterTop
    End With

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cStatusBar.StatusBarControl.Let", eGDRaiseError_Raise

End Property

Property Let Steps(pData As Long)
On Error GoTo ErrSection:

    m.lStep = 0
    m.lSteps = pData

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cStatusBar.Steps.Let", eGDRaiseError_Raise

End Property

Property Let Message(pData As String)
On Error GoTo ErrSection:

    m.StatusBar.TextMatrix(0, 0) = pData
    If pData <> "" Then
        m.bUseMessage = True
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cStatusBar.Message.Let", eGDRaiseError_Raise

End Property

Property Get Steps() As Long
On Error GoTo ErrSection:

    Steps = m.lSteps

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cStatusBar.Steps.Get", eGDRaiseError_Raise

End Property

Public Sub UpdateMeter(Optional pNewIncrementValue As Variant = 0)
On Error GoTo ErrSection:

    If m.lSteps = 0 Then
        m.lSteps = 100
    End If
    If pNewIncrementValue > 0 Then
        m.lStep = pNewIncrementValue
    Else
        m.lStep = m.lStep + 1
    End If
    
    If gdTickCount - m.dLastGridRefresh > 500 Then
        With m.StatusBar
            .Redraw = False
            If m.lStep < m.lSteps Then
                If Not m.bUseMessage Then
                    .TextMatrix(0, 0) = Format(m.lStep / m.lSteps, "#,##0%")
                End If
                .Cell(flexcpFloodPercent, 0, 0) = Int((m.lStep / m.lSteps) * 100)
            Else
                If Not m.bUseMessage Then
                    .TextMatrix(0, 0) = Format(1, "#,##0%")
                End If
                .Cell(flexcpFloodPercent, 0, 0) = 100
            End If
            .Row = 0
            .Col = 0
            .Redraw = True
            .Refresh
        End With
        m.dLastGridRefresh = gdTickCount
    Else
        With m.StatusBar
            .Redraw = False
            If m.lStep >= m.lSteps Then
                If Not m.bUseMessage Then
                    m.StatusBar.TextMatrix(0, 0) = Format(1, "#,##0%")
                End If
                m.StatusBar.Cell(flexcpFloodPercent, 0, 0) = 100
            End If
            .Redraw = True
            .Refresh
        End With
    End If
    m.bUseMessage = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStatusBar.UpdateMeter", eGDRaiseError_Raise

End Sub

Property Let Value(pData As Long)
On Error GoTo ErrSection:

    With m.StatusBar
        .Redraw = False
        .Cell(flexcpFloodPercent, 0, 0) = pData
        .TextMatrix(0, 0) = Format(pData, "#0") & "%"
        .Cell(flexcpFontBold, 0, 0) = True
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterTop
        .Redraw = True
        .Refresh
    End With

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cStatusBar.Value.Let", eGDRaiseError_Raise

End Property

Private Sub Class_Terminate()
    Set m.StatusBar = Nothing
End Sub
