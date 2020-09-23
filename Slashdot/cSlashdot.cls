VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSlashdot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************
'* Class Name:  cSlashdot
'* Author:      Jeremy Howard aka FishOutaWata - www.fishoutawata.com
'* Date:        3/23/2001
'* Description: Responsible for connecting to Slashdot.org and retrieving
'*              Articles
'**************************************************************************

Private objArticles As cArticles
Private iCount      As Integer

Public Function RetieveArticles() As Boolean

On Error GoTo ERR_HANDLER

Dim objXML        As New MSXML.XMLHTTPRequest
Dim objDOM        As New MSXML.DOMDocument
Dim objNodeList   As MSXML.IXMLDOMNodeList
Dim objNode       As MSXML.IXMLDOMNode
Dim objRoot       As MSXML.IXMLDOMElement
Dim objTitleNode  As MSXML.IXMLDOMNode
Dim objUrlNode    As MSXML.IXMLDOMNode
Dim objTimeNode   As MSXML.IXMLDOMNode
Dim objAuthorNode As MSXML.IXMLDOMNode

Dim bRetVal       As Boolean

'Connect to Slashdot.org and retieve slashdot.xml
objXML.open "GET", "http://slashdot.org/slashdot.xml", False
objXML.setRequestHeader "pragma", "no-cache"
objXML.setRequestHeader "cache-control", "no-cache"
objXML.send

objDOM.async = False

bRetVal = objDOM.Load(objXML.responseXML)

'Make sure the document loaded
If bRetVal = True Then
    'Parse XML for data needed
    Set objRoot = objDOM.documentElement
    Set objNodeList = objRoot.selectNodes("story")
    
    'get the Article count for Count Property
    iCount = objNodeList.length
    
    'Iterate through NodeList
    For Each objNode In objNodeList
        Set objTitleNode = objNode.selectSingleNode("title")
        Set objUrlNode = objNode.selectSingleNode("url")
        Set objTimeNode = objNode.selectSingleNode("time")
        Set objAuthorNode = objNode.selectSingleNode("author")
        
        'Add Article to Articles object
        objArticles.Add objTitleNode.Text, objAuthorNode.Text, objTimeNode.Text, objUrlNode.Text
    Next
    
    RetieveArticles = True
    GoTo CLEAN_UP
Else
    RetieveArticles = False
    GoTo CLEAN_UP
End If

ERR_HANDLER:
    Err.Raise Err.Number

CLEAN_UP:
    Set objXML = Nothing
    Set objDOM = Nothing
    Set objNodeList = Nothing
    Set objNode = Nothing
    Set objRoot = Nothing
    Set objTitleNode = Nothing
    Set objUrlNode = Nothing
    Set objTimeNode = Nothing
    Set objAuthorNode = Nothing
    
    
End Function

Public Property Get ArticleCount() As Integer

ArticleCount = iCount

End Property


Public Function GetArticle(iIndex As Integer) As cArticle

Set GetArticle = objArticles.Item(iIndex)

End Function

Private Sub Class_Initialize()

Set objArticles = New cArticles

End Sub


Private Sub Class_Terminate()

Set objArticles = Nothing

End Sub


