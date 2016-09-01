Attribute VB_Name = "mSubClasser"
'******************************************************
'ezSubClasser.bas
'Example of one method of native subclassing
'by Ray Mercer <raymer@macnica.co.jp>
'Copyright (C) 1999 by Ray Mercer, All rights reserved
'http://i.am/shrinkwrapvb
'*******************************************************
'Updated 02/25/99, Ray Mercer
'Changed entire subclassing technique to my original procedure
'Updated 07/28/99, Ray Mercer
'Revised entire module to make it useful as a generic subclassing tool
    'This technique is based on the one used by Bruce McKinney
    'in "Hardcore VB" - but gone through several permutations
    'now simplified and very generic - I hope :-)

'Thanks Ray for this subclasser, it's the best!
'ProcModule thing added by Thomas Allin, vemod@musiker.nu, April 2000.

Option Private Module
Option Explicit

Private Function PropName(ByVal hWnd As Long) As String
    PropName = "mSubClasser:" & Str(hWnd)
End Function

Public Function ezSubclass(ByVal hWnd As Long, ByRef ProcModule As cWindowLink) As Long 'returns procOld
    
    Dim ProcOld As Long
    
    If SubclassingEnabled Then
        ' attach object pointer to window
        ' (need to do this since VB does not handle callbacks in class/form)
        Call SetProp(hWnd, PropName(hWnd), ObjPtr(ProcModule))
        
        ' Subclass window by installing window procedure
        ProcOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndCallBackProcDirector)
        If ProcOld = 0 Then
            Debug.Print "ERROR! Cannot subclass this window!"
        End If
    End If

    ezSubclass = ProcOld

End Function

Public Sub ezUnSubclass(ByVal hWnd As Long, ByVal OldProc As Long)
        
    If OldProc Then
        'reclaim resources
        Call RemoveProp(hWnd, PropName(hWnd))
        
        'Unsubclass by reassigning old window procedure
        Call SetWindowLong(hWnd, GWL_WNDPROC, OldProc)
        OldProc = 0
    End If
    
End Sub

Function WndCallBackProcDirector(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    'Debug.Print "ENTERED CALLBACK ROUTER"
    Dim pObj As Long
    Dim objToCall As cWindowLink
    Dim objTemp As cWindowLink
    
    ' Get the object pointer for the current OCX/Form/Class instance from the window handle
    pObj = GetProp(hWnd, PropName(hWnd))
    'Debug.Assert pObj <> 0
    If 0 = pObj Then
        Debug.Print "CRITICAL SUBCLASSING ERROR!"
        'TODO add some cleanup?
        Exit Function
    End If
   
    ' Turn the pointer into an illegal, uncounted interface
    CopyMemory objTemp, pObj, 4
    ' Assign to legal reference
    Set objToCall = objTemp
    ' Destroy the illegal reference
    CopyMemory objTemp, 0&, 4
    
    ' Use the interface to call back to the class
    WndCallBackProcDirector = objToCall.WindowProc(hWnd, iMsg, wParam, lParam)
    Set objToCall = Nothing

End Function

Public Function SubclassingEnabled() As Boolean
    
    Static iEnableSubclass As Integer
    
    If iEnableSubclass = 0 Then
        If FileExist(App.Path & "\SkipWindowLink.ALL") Then
            iEnableSubclass = -1
        ElseIf IsIDE And FileExist(App.Path & "\SkipWindowLink.IDE") Then
            iEnableSubclass = -1
        Else
            iEnableSubclass = 1
        End If
    End If
    
    SubclassingEnabled = (iEnableSubclass >= 0)

End Function

