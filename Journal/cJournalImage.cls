VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJournalImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cJournalImage.cls
'' Description: Class for handling journal image information
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lID As Long                         ' Unique Identifier for the journal image
    nImageType As eGDJournalImageTypes  ' Type of the image (chart, order form, etc.)
    strFileName As String               ' Filename for the image
    strCaption As String                ' Caption for the image
    lOrderJournalID As Long             ' ID for the corresponding order journal
    lDateJournalID As Long              ' ID for the corresponding date journal
End Type
Private m As mPrivate

Public Property Get ID() As Long
    ID = m.lID
End Property
Public Property Let ID(ByVal lID As Long)
    m.lID = lID
End Property

Public Property Get ImageType() As eGDJournalImageTypes
    ImageType = m.nImageType
End Property
Public Property Let ImageType(ByVal nImageType As eGDJournalImageTypes)
    m.nImageType = nImageType
End Property

Public Property Get FileName() As String
    FileName = m.strFileName
End Property
Public Property Let FileName(ByVal strFileName As String)
    m.strFileName = strFileName
End Property

Public Property Get Caption() As String
    Caption = m.strCaption
End Property
Public Property Let Caption(ByVal strCaption As String)
    m.strCaption = strCaption
End Property

Public Property Get OrderJournalID() As Long
    OrderJournalID = m.lOrderJournalID
End Property
Public Property Let OrderJournalID(ByVal lOrderJournalID As Long)
    m.lOrderJournalID = lOrderJournalID
End Property

Public Property Get DateJournalID() As Long
    DateJournalID = m.lDateJournalID
End Property
Public Property Let DateJournalID(ByVal lDateJournalID As Long)
    m.lDateJournalID = lDateJournalID
End Property
