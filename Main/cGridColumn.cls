VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGridColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cGridColumn.cls
'' Description: Meta information about a grid column
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kNumFields As Long = 6&

Private Type mPrivate
    bVisible As Boolean                 ' Is the column visible?
    strName As String                   ' Name of the column
    lLevel As Long                      ' Level of the column
    lPosition As Long                   ' Position of the column
    lWidth As Long                      ' Width of the column
    bShowInSettings As Boolean          ' Show the column in the settings form?
End Type
Private m As mPrivate

Public Property Get Visible() As Boolean
    Visible = m.bVisible
End Property
Public Property Let Visible(ByVal bVisible As Boolean)
    m.bVisible = bVisible
End Property

Public Property Get Name() As String
    Name = m.strName
End Property
Public Property Let Name(ByVal strName As String)
    m.strName = strName
End Property

Public Property Get Level() As Long
    Level = m.lLevel
End Property
Public Property Let Level(ByVal lLevel As Long)
    m.lLevel = lLevel
End Property

Public Property Get Position() As Long
    Position = m.lPosition
End Property
Public Property Let Position(ByVal lPosition As Long)
    m.lPosition = lPosition
End Property

Public Property Get Width() As Long
    Width = m.lWidth
End Property
Public Property Let Width(ByVal lWidth As Long)
    m.lWidth = lWidth
End Property

Public Property Get ShowInSettings() As Boolean
    ShowInSettings = m.bShowInSettings
End Property
Public Property Let ShowInSettings(ByVal bShowInSettings As Boolean)
    m.bShowInSettings = bShowInSettings
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Fill in the member variables from a settings string
'' Inputs:      Settings String
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strSettings As String)
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray      ' Array of fields from the delimited string
    
    astrFields.SplitFields strSettings, ";"
    
    m.bVisible = (astrFields(0) = "1")
    m.strName = astrFields(1)
    m.lLevel = CLng(Val(astrFields(2)))
    m.lPosition = CLng(Val(astrFields(3)))
    m.lWidth = CLng(Val(astrFields(4)))
    If Len(astrFields(5)) = 0 Then m.bShowInSettings = False Else m.bShowInSettings = (astrFields(5) = "1")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cGridColumn.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Create a settings string from the member variables
'' Inputs:      None
'' Returns:     Settings String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray      ' Array of fields from the delimited string
    
    astrFields.Create eGDARRAY_Strings, kNumFields
    
    If m.bVisible Then astrFields(0) = "1" Else astrFields(0) = "0"
    astrFields(1) = m.strName
    astrFields(2) = Str(m.lLevel)
    astrFields(3) = Str(m.lPosition)
    astrFields(4) = Str(m.lWidth)
    If m.bShowInSettings Then astrFields(5) = "1" Else astrFields(5) = "0"
    
    ToString = astrFields.JoinFields(";")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cGridColumn.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    m.bVisible = True
    m.strName = ""
    m.lLevel = 1&
    m.lPosition = 0&
    m.lWidth = 0&
    m.bShowInSettings = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cGridColumn.Class_Initialize"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up class members when object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cGridColumn.Class_Terminate"
    
End Sub

