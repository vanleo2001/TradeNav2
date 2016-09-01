VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' cForms -- written by TLB 11/26/2013
'
' PROBLEM:  The reason this class got created was to solve the issues that we've realized can
' and do happen when using either a "For i = 0 to Forms.Count - 1" or "For Each frm In Forms" loop.
' The problems when using an index of the Forms collection, e.g. Forms(i), are that the collection can
' change while your loop is processing -- e.g. if a form is removed from the collection while your
' loop is in progress, a form can get skipped, or even worse the form at a specific index can suddenly
' be referencing a different form than it was, or it can cause the For loop to run past the newer .Count
' and cause an "invalid subscript" error.  And the problem with the "For Each frm ..." iteration is that
' it appears you get some kind of copy of the Forms collection -- then if a form gets removed while this
' loop is in progress, a reference to that form sticks around and never ends up completely going away
' (even after the iterator loop has been exited!).
'
' SOLUTION:  This class creates an array of window handles the first time .NextWindow is called.  Since
' no actual "pointer" references to the forms are kept, it allows a form to be completely removed from
' the Forms collection even while the loop using .NextWindow is still processing (except of course for
' the window which is currently being "used" in the loop, since a reference only to it was returned).
' And subsequent calls to .NextForm always look through the current Forms collection for the window that
' was next in the original array (and if it no longer exists, then it looks for the next window, etc.).
' So the .NextWindow call simply returns the next Form that still exists from what had existed at the
' time of the first call, until no other forms match and it returns "Nothing".
'
#If 0 Then
' USAGE (see 2 examples below):
' 1. create a new cForms instance (Dim frms as New cForms)
' 2. optionally call .Init prior to the loop (e.g. if wanting to filter for only a specific type of form,
'       or required if restarting another loop after it had already been used previously in that routine)
' 3. run a "Do ... Loop" that keeps calling .NextForm, and exit the Do Loop when .NextForm returns Nothing.
'
' #1) Example of generic usage (when checking all types of forms):
Private Sub UsageExample_GenericForms()

    Dim frm As Form, frms As New cForms

    frms.Init ' (optional in this case, but probably a good practice)
    Do
        Set frm = frms.NextForm
        If frm Is Nothing Then Exit Do
        ' then simply use the "frm" returned by .NextForm
        frm.Tag = "Hello"
    Loop
    Set frm = Nothing

End Sub
'
' #2) Example of filtering for only a specific type of form:
Private Sub UsageExample_SpecificForms()

    Dim frm As frmTickDistribution, frms As New cForms

    frms.Init frmTickDistribution ' can specify the only type of form that should be returned
    Do
        Set frm = frms.NextForm
        If frm Is Nothing Then Exit Do
        ' then simply use the "frm" returned by .NextForm
        frm.ReleaseMarketDepth
    Loop
    Set frm = Nothing

End Sub
#End If

'===================================== cForms class =====================================================
' stuff used internally:
Private Type mPrivate
    strFormType As String   ' if filtering for a specific type of form
    aHwnds() As Long        ' to store the window handles of the existing Forms collection when first called
    iNext As Long           ' an index to increment through the array of window handles
End Type
Private m As mPrivate

Private Sub Class_Initialize()
    Init
End Sub

Private Sub Class_Terminate()
    Init
End Sub

' can optionally pass in a specific type of Form to filter on (so will only return forms of that type)
Public Sub Init(Optional TypeOfForm As Form = Nothing)

    ' clear things
    m.iNext = -1
    ReDim aHwnds(0)
    ' if filtering for a specific type of form, store it's TypeName
    If TypeOfForm Is Nothing Then
        m.strFormType = ""
    Else
        m.strFormType = TypeName(TypeOfForm)
    End If

End Sub

Public Function NextForm() As Form

    Dim i&, hWnd&
    Dim frm As Form

    If Not IsIDE Then
        On Error Resume Next
    End If
    
    ' if first call after the .Init, then create the array of window handles for the current Forms collection
    If m.iNext < 0 Then
        GetForms
        m.iNext = -1
    End If
    
    ' get the next window handle in the array
    Do While m.iNext < UBound(m.aHwnds)
        m.iNext = m.iNext + 1
        hWnd = m.aHwnds(m.iNext)
        If hWnd <> 0 Then ' skip 0's (either no longer a form or not the right type)
            ' look for a form in the current Forms collection with a matching window handle
            For i = 0 To Forms.Count - 1
                ' always verify still in-bounds (in case Forms collection has changed size)
                If i >= Forms.Count Then Exit For
                Set frm = Forms(i)
                If frm.hWnd = hWnd Then
                    Set NextForm = frm ' return the matching Form
                    Exit Do
                End If
            Next
        End If
    Loop ' if that form no longer exists, then loop back to check for the next window handle in the array
    Set frm = Nothing

End Function

' this routine used when .NextWindow is first called to get array of window handles currently in the Forms collection
Private Sub GetForms()
    
    Dim i&, hWnd&
    Dim frm As Form
    
    If Not IsIDE Then
        On Error Resume Next
    End If
    
    ReDim m.aHwnds(Forms.Count) ' size should be one item bigger than what we really need
    For i = 0 To UBound(m.aHwnds)
        hWnd = 0
        ' always verify still in-bounds (in case Forms collection has changed size)
        If i < Forms.Count Then
            Set frm = Forms(i)
            If Not frm Is Nothing Then
                ' check if filtering for a specific type of form
                If Len(m.strFormType) = 0 Then
                    hWnd = frm.hWnd
                ElseIf UCase(TypeName(frm)) = UCase(m.strFormType) Then
                    hWnd = frm.hWnd
                End If
            End If
            Set frm = Nothing
        End If
        ' store window handle (or 0 if no form or not the right type of form)
        m.aHwnds(i) = hWnd
    Next
    
End Sub

