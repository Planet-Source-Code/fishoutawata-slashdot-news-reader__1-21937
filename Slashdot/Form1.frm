VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   1740
      TabIndex        =   0
      Top             =   1350
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim objXML As New MSXML.XMLHTTPRequest
Dim objDOM As New MSXML.DOMDocument
Dim objDOMList As MSXML.IXMLDOMNodeList
Dim objNode As MSXML.IXMLDOMNode
Dim objRoot As MSXML.IXMLDOMElement
Dim objTransactionNode As MSXML.IXMLDOMNode
Dim objAttribNode As MSXML.IXMLDOMNode
Dim objAttribs As MSXML.IXMLDOMNamedNodeMap

objXML.open "GET", "http://slashdot.org/slashdot.xml", False
objXML.setRequestHeader "pragma", "no-cache"
objXML.send

objDOM.async = False
bretval = objDOM.Load(objXML.responseXML)

Debug.Print bretval

Set objRoot = objDOM.documentElement
Set objDOMList = objRoot.selectNodes("story")

Set objTest = objRoot.selectSingleNode("story")

Debug.Print objTest.childNodes.Item(0).Text


Debug.Print objDOMList.length

For x = 1 To objDOMList.length
    Set objNode = objDOMList.Item(0)
    Debug.Print objNode.childNodes.Item(0).Text
Next

'For Each objNode In objDOMList
'    For x = 0 To objNode.childNodes.length - 1
'        Debug.Print objNode.childNodes.Item(x).Text
'    Next
'Next

End Sub
