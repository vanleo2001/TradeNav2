VERSION 5.00
Begin VB.Form frmDisplayImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgPicture 
      Height          =   3015
      Left            =   120
      Top             =   120
      Width           =   4515
   End
End
Attribute VB_Name = "frmDisplayImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmDisplayImage.frm
'' Description: Form that displays an image
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 09/08/2014   DAJ         Getting form icon different way
'' 10/24/2014   DAJ         Core Application functions for DLL's
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal strCaption As String, ByVal strImageFilename As String)
On Error GoTo ErrSection:

    Icon = g.AppBridge.OwnerFormIcon
    Caption = strCaption
    imgPicture.Picture = LoadPicture(strImageFilename)
    SizeForm
    
    If g.bAppIsIde Then
        mGenesis.ShowForm Me, eForm_Nonmodal
    Else
        g.TnCore.ShowForm Me, eForm_Nonmodal
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDisplayImage.ShowMe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SizeForm
'' Description: Size the form based on the size of the image
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SizeForm()
On Error GoTo ErrSection:

    Dim lHeightDiff As Long             ' Difference between the Height and the ScaleHeight
    Dim lWidthDiff As Long              ' Difference between the Width and the ScaleWidth
    
    lHeightDiff = Height - ScaleHeight
    lWidthDiff = Width - ScaleWidth
    
    Move Left, Top, imgPicture.Width + 240 + lWidthDiff, imgPicture.Height + 240 + lHeightDiff

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "frmDisplayImage.SizeForm"
    
End Sub

Private Sub Form_Load()
    g.Styler.StyleForm Me
End Sub
