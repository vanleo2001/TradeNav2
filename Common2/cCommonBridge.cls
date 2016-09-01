VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCommonBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'This class is used as a bridge between the Main application and this one.
'It is passed references to objects created in the caller such as Systems
'and database variables to Ms-Acess.

Property Let dbNavRef(pData As Database)
    Set g.dbNav = pData
    RefreshEditorOptions
End Property

Public Sub RefreshEditorOptions()
    Set g.EditorOptions = New cEditorOptions
    g.EditorOptions.Load
End Sub

Property Let AppPath(ByVal strAppPath As String)
    g.strAppPath = strAppPath
End Property

Property Get AppPath() As String
    AppPath = g.strAppPath
End Property

Property Let CustomerID(ByVal lLCD As Long)
    g.lLCD = lLCD
End Property

Property Get CustomerID() As Long
    CustomerID = g.lLCD
End Property
