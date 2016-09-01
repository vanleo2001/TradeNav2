VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCoreBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cCoreBridge.cls
'' Description: Bridge for the Core DLL
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
''
'' 10-01-2015   J. Muchmore added support for TN toolbar chrome vs classic images
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    CoreForm As frmCore                 ' Core form
End Type
Private m As mPrivate

Public Function Image16() As Object
    Set Image16 = frmCore.img16
End Function

Public Function ImageList1() As Object
    Set ImageList1 = frmCore.ImageList1
End Function

Public Function ImageList2() As Object
    Set ImageList2 = frmCore.ImageList2
End Function

Public Function ImageListToolbar() As Object
    'JM 10-15-2015: obsolete, use ImgListToolbarExt instead
    Set ImageListToolbar = frmCore.imglistClassic16
End Function

Public Function ImageListToolbar32() As Object
    'JM 10-15-2015: obsolete, should use ImgListToolbarExt instead
    Set ImageListToolbar32 = frmCore.imglistClassic32
End Function

Public Function ImageListMouse22() As Object
    Set ImageListMouse22 = frmCore.ImgListMouse22
End Function

Public Function ImageListMouse35() As Object
    Set ImageListMouse35 = frmCore.imgListMouse35
End Function

Public Function ImageListMouse45() As Object
    Set ImageListMouse45 = frmCore.imgListMouse45
End Function

Public Function ImageListMouse59() As Object
    Set ImageListMouse59 = frmCore.ImgListMouse59
End Function

Public Function ImageListMouse110() As Object
    Set ImageListMouse110 = frmCore.imgListMouse110
End Function

Public Function ImageListMouse118() As Object
    Set ImageListMouse118 = frmCore.imgListMouse118
End Function

Public Function ImageListBackgroundSkin() As Object
    Set ImageListBackgroundSkin = frmCore.imgListBkSkin
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Picture16
'' Description: Return the appropriate icon
'' Inputs:      Picture Name, Image List
'' Returns:     Icon
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Picture16(ByVal strPicture As String, Optional ByVal iImageList As Integer = 0) As Object
On Error Resume Next

    Dim i&

    If (InStr(strPicture, "Theme=") <> 0) Then
        If InStr(strPicture, "Classic") <> 0 Then
            For i = frmCore.imglistChromeLight16.ListImages.Count To 1 Step -1
                frmCore.imglistChromeLight16.ListImages.Remove i
            Next
            For i = frmCore.imglistChromeLight32.ListImages.Count To 1 Step -1
                frmCore.imglistChromeLight32.ListImages.Remove i
            Next
            For i = frmCore.imglistChromeDark16.ListImages.Count To 1 Step -1
                frmCore.imglistChromeDark16.ListImages.Remove i
            Next
            For i = frmCore.imglistChromeDark32.ListImages.Count To 1 Step -1
                frmCore.imglistChromeDark32.ListImages.Remove i
            Next
            For i = frmCore.imglistNoMask.ListImages.Count - 1 To 1 Step -1
                frmCore.imglistNoMask.ListImages.Remove i
            Next
        ElseIf InStr(strPicture, "Charcoal") <> 0 Then
            For i = frmCore.imglistChromeDark16.ListImages.Count To 1 Step -1
                If frmCore.imglistChromeDark16.ListImages(i).Key <> "kDollar" Then
                    frmCore.imglistChromeDark16.ListImages.Remove i
                End If
            Next
            For i = frmCore.imglistChromeDark32.ListImages.Count To 1 Step -1
                If frmCore.imglistChromeDark32.ListImages(i).Key <> "kDollar" Then
                    frmCore.imglistChromeDark32.ListImages.Remove i
                End If
            Next
            For i = frmCore.imglistClassic16.ListImages.Count To 1 Step -1
                frmCore.imglistClassic16.ListImages.Remove i
            Next
            For i = frmCore.imglistClassic32.ListImages.Count To 1 Step -1
                frmCore.imglistClassic32.ListImages.Remove i
            Next
            For i = frmCore.imglistNoMaskClassic.ListImages.Count To 1 Step -1
                frmCore.imglistNoMaskClassic.ListImages.Remove i
            Next
        ElseIf InStr(strPicture, "Ivory") <> 0 Then
            For i = frmCore.imglistChromeLight16.ListImages.Count To 1 Step -1
                frmCore.imglistChromeLight16.ListImages.Remove i
            Next
            For i = frmCore.imglistChromeLight32.ListImages.Count To 1 Step -1
                frmCore.imglistChromeLight32.ListImages.Remove i
            Next
            For i = frmCore.imglistClassic16.ListImages.Count To 1 Step -1
                frmCore.imglistClassic16.ListImages.Remove i
            Next
            For i = frmCore.imglistClassic32.ListImages.Count To 1 Step -1
                frmCore.imglistClassic32.ListImages.Remove i
            Next
            For i = frmCore.imglistNoMaskClassic.ListImages.Count To 1 Step -1
                frmCore.imglistNoMaskClassic.ListImages.Remove i
            Next
        End If
        Set Picture16 = frmCore.img16.ListImages("kBlank").Picture
        Exit Function
    End If
    
    
    Select Case Abs(iImageList)
        Case 0:
            'JM 01-27-2016: The default imagelist is img16 which originally contained all classic
            '   style icons. Trade Navigator Theme feature required 3 sets of icons. To minimize
            '   gdi objects count, icons from themes that are not in use get deleted from the
            '   appropriate image lists right at start-up. img16 now contains only icons that
            '   are shared by all themes. Classic icons that can be deleted when the current
            '   theme is Charcoal or Ivory have been moved to imglistNoMaskClassic and imglistClassic16.
            '   When an image is requested from img16, we now check the other 2 image lists before
            '   checking the original img16 imagelist.
            Set Picture16 = frmCore.imglistNoMaskClassic.ListImages(strPicture)
            If Not Picture16 Is Nothing Then
                Set Picture16 = frmCore.imglistNoMaskClassic.ListImages(strPicture).Picture
            End If
        Case 1:
            Set Picture16 = frmCore.ImageList1.ListImages(strPicture).ExtractIcon
        Case 2:
            Set Picture16 = frmCore.ImageList2.ListImages(strPicture).Picture
        Case 3:
            Set Picture16 = frmCore.imglistChromeLight16.ListImages(strPicture)
            If Not Picture16 Is Nothing Then
                Set Picture16 = frmCore.imglistChromeLight16.ListImages(strPicture).Picture
            End If
        Case 4:
            Set Picture16 = frmCore.imglistChromeDark16.ListImages(strPicture)
            If Not Picture16 Is Nothing Then
                Set Picture16 = frmCore.imglistChromeDark16.ListImages(strPicture).Picture
            End If
        Case 5:
            Set Picture16 = frmCore.imglistNoMask.ListImages(strPicture)
            If Not Picture16 Is Nothing Then
                Set Picture16 = frmCore.imglistNoMask.ListImages(strPicture).Picture
            End If
    End Select
    
    If Picture16 Is Nothing Then
        Set Picture16 = frmCore.img16.ListImages(strPicture)
        If Picture16 Is Nothing Then
            If iImageList = 0 Then
                Set Picture16 = frmCore.imglistClassic16.ListImages(strPicture)
            End If
            If Picture16 Is Nothing Then
                Set Picture16 = frmCore.img16.ListImages("kBlank").Picture
            Else
                Set Picture16 = frmCore.imglistClassic16.ListImages(strPicture).Picture
            End If
        Else
            Set Picture16 = frmCore.img16.ListImages(strPicture).Picture
        End If
    End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.CoreForm = New frmCore

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCoreBridge.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.CoreForm = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCoreBridge.Class_Terminate"
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImgListToolbarExt (stands for imglisttoolbarExtend)
'' Description: Selects requested image list for Trade Navigator main tool bars
'' Inputs:      strImgStyle: Classic, Dark, Light
''              strImgKey: key into imagelist collection
''              strToolbar: currently not used, can send empty string
''              imgSize: 16, 32
'' Returns:     ListImage object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ImgListToolbarExt(ByVal strImgStyle As String, ByVal strImgKey As String, _
    ByVal strToolbar As String, ByVal imgSize As Long) As Object
    
    On Error Resume Next

    If UCase(strImgStyle) = "DARK" Then
        If imgSize = 32 Then
            Set ImgListToolbarExt = frmCore.imglistChromeDark32.ListImages(strImgKey)
        Else
            Set ImgListToolbarExt = frmCore.imglistChromeDark16.ListImages(strImgKey)
        End If
    ElseIf UCase(strImgStyle) = "LIGHT" Then
        If imgSize = 32 Then
            Set ImgListToolbarExt = frmCore.imglistChromeLight32.ListImages(strImgKey)
        Else
            Set ImgListToolbarExt = frmCore.imglistChromeLight16.ListImages(strImgKey)
        End If
    ElseIf imgSize = 32 Then
        Set ImgListToolbarExt = frmCore.imglistClassic32.ListImages(strImgKey)
    Else
        Set ImgListToolbarExt = frmCore.imglistClassic16.ListImages(strImgKey)
    End If
    
End Function

