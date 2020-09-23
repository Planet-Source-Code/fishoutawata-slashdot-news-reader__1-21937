VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
'* Class Name:   cArticle
'* Author:       Jeremy Howard aka FishOutaWata - www.fishoutawata.com
'* Date:         3/23/2001
'* Description:  Contains details on one paticular Slashdot Article
'*************************************************************************

Private sTitle     As String
Private sAuthor    As String
Private sSubmitted As String
Private sUrl       As String

Public Property Get Title() As String

Title = sTitle

End Property

Public Property Let Title(ByVal sNewValue As String)

sTitle = sNewValue

End Property

Public Property Get Author() As String

Author = sAuthor

End Property

Public Property Let Author(ByVal sNewValue As String)

sAuthor = sNewValue

End Property

Public Property Get Submitted() As String

Submitted = sSubmitted
End Property

Public Property Let Submitted(ByVal sNewValue As String)

sSubmitted = sNewValue

End Property

Public Property Get Url() As String

Url = sUrl

End Property

Public Property Let Url(ByVal sNewValue As String)

sUrl = sNewValue

End Property
