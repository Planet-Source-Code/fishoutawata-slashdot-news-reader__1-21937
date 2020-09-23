VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArticles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
'* Class Name:  cArticles
'* Author:      Jeremy Howard aka FishOutaWata - www.fishoutawata.com
'* Date:        3/23/2001
'* Description: Maintains the collection of cArticle objects
'*************************************************************************

Private m_colArticles As Collection


Private Sub Class_Initialize()

'Create the collection object
Set m_colArticles = New Collection

End Sub

Private Sub Class_Terminate()

'Destroy the collection object
Set m_colArticles = Nothing

End Sub

Public Sub Add(ByVal sTitle As String, ByVal sAuthor As String, _
               ByVal sSubmitted As String, ByVal sUrl As String)

Dim objArticle As cArticle

'Create new Article Object
Set objArticle = New cArticle

With objArticle
    'Populate Article object
    .Title = Trim(sTitle)
    .Author = Trim(sAuthor)
    .Submitted = Trim(sSubmitted)
    .Url = Trim(sUrl)
End With

'Add Article object to collection
m_colArticles.Add objArticle

'Destroy Article Object
Set objArticle = Nothing

End Sub

Public Sub Clear()

Dim x As Integer

'Clear the Collection Object
For x = 1 To m_colArticles.Count
    m_colArticles.Remove (x)
Next

End Sub

Public Function Item(iIndex) As cArticle

Set Item = m_colArticles.Item(iIndex)

End Function
