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
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cStatusBar.cls
'' Description: Make a VideoSoft grid into a status bar
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    fgStatusBar As Object               ' VideoSoft grid
End Type
Private m As mPrivate

Public Property Let StatusBar(pData As Object)
    Set m.fgStatusBar = pData
End Property

Public Property Get StatusBar() As Object
    StatusBar = m.fgStatusBar
End Property

Public Property Let Value(pData As Long)
    
    With m.fgStatusBar
        .Redraw = False
        .Cell(flexcpBackColor, 0, 0) = &H808080   'vbButtonFace
        .Cell(flexcpForeColor, 0, 0) = &HFFFFFF   'vbBlack
        .Cell(flexcpFloodColor, 0, 0) = &HFF0000
        
        .Cell(flexcpFloodPercent, 0, 0) = pData
        .TextMatrix(0, 0) = Format(pData / 100, "#0%")
        .Cell(flexcpFontBold, 0, 0) = True
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterTop
        .Redraw = True
        .Refresh
    End With

End Property

Public Property Get Value() As Long
    Value = m.fgStatusBar.Cell(flexcpFloodPercent, 0, 0)
End Property

Private Sub Class_Terminate()
    Set m.fgStatusBar = Nothing
End Sub

Public Sub Refresh()
    m.fgStatusBar.Refresh
End Sub
