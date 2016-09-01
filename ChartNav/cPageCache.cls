VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPageCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    Tree As cGdTree                 'tree of cChartCache objects belonging to this page
    idxActiveChart As Long          'tree index of chart object that is the active chart
    strPageName As String
    bMaximized As Boolean
End Type

Private m As mPrivate

Private Sub Class_Initialize()
    m.idxActiveChart = -1
    m.strPageName = ""
End Sub

Private Sub Class_Terminate()
    
    On Error Resume Next
    
    m.Tree.Clear
    Set m.Tree = Nothing

End Sub

Public Property Get PageName() As String
    PageName = m.strPageName
End Property

Public Property Let PageName(ByVal strName$)
    m.strPageName = strName
End Property

Public Property Get ChartsCount() As Long
On Error GoTo ErrSection:

    If Not m.Tree Is Nothing Then
        ChartsCount = m.Tree.Count
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPageCache.ChartsCount.Get"

End Property

Public Property Get CachedObject(ByVal idx&) As cChartCache
On Error GoTo ErrSection:
    
    If Not m.Tree Is Nothing Then
        Set CachedObject = m.Tree(idx)
    End If

ErrExit:
    Exit Property

ErrSection:
    RaiseError "cPageCache.CachedObject.Get"

End Property

Public Property Get ChartsMaximized() As Boolean
    ChartsMaximized = m.bMaximized
End Property

Public Property Get ActiveChartIdx() As Long
    ActiveChartIdx = m.idxActiveChart
End Property

Public Sub ActiveIdxReset()
    m.idxActiveChart = -1
End Sub

Public Sub MoveActiveToEnd(ByVal bMaximized As Boolean)
On Error GoTo ErrSection:
        
    Dim i&, Cached As cChartCache
    
    m.bMaximized = bMaximized
    
    If Not m.Tree Is Nothing Then
        If m.idxActiveChart > 0 And m.idxActiveChart <= m.Tree.Count Then
            Set Cached = m.Tree(m.idxActiveChart)
            m.Tree.Remove m.idxActiveChart
            i = m.Tree.Add(Cached)
            If i = m.Tree.Count Then
                m.idxActiveChart = i
            Else
                'theoretically should never get here
                m.idxActiveChart = m.Tree.Count
                DebugLog "Error moving cached chart object to end of tree: " & m.strPageName & " (i=" & Str(i) & ", count=" & Str(m.Tree.Count)
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPageCache.MoveActiveToEnd"

End Sub

Public Function PageCacheAdd(Chart As cChart) As Boolean
On Error GoTo ErrSection:

    Dim frm As Form
    Dim ChartCache As cChartCache
    Dim bSaveIndex As Boolean

    If Chart Is Nothing Then Exit Function
    If Chart.Form Is Nothing Then Exit Function
    
    Set frm = Chart.Form
    If frm.IsInGameMode Then
        PageCacheAdd = True           'return success as we don't restore game mode charts anyhow
        Exit Function
    End If
    
    If frm Is g.ChartGlobals.frmActiveNonDetached Then bSaveIndex = True
    
    Set ChartCache = New cChartCache
    ChartCache.CacheChartSave Chart
    
    If m.Tree Is Nothing Then Set m.Tree = New cGdTree
    m.Tree.Add ChartCache
    
    If bSaveIndex Then m.idxActiveChart = m.Tree.Count
    
    PageCacheAdd = True

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cPageCache.PageCacheAdd"

End Function

Public Sub PageCacheReleaseObjects()
On Error GoTo ErrSection:

    Dim i&
    
    'release chart objects without destroying them so they can be saved to new Page object
    If Not m.Tree Is Nothing Then
        For i = 1 To m.Tree.Count
            If Not m.Tree(i) Is Nothing Then m.Tree(i).CacheChartReleaseObj
        Next
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPageCache.PageCacheReleaseObjects"

End Sub
