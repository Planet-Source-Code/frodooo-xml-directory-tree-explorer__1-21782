VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Private mCurrKey As String
Private mXMLIsLoaded As Boolean
Private mCurrFileName As String
Private oDoc As DOMDocument
Private mTreeView As TreeView
Private mMaxID As Integer

Public Sub ChangeAttribute(AttributeName As String, NewValue As String)
'changes an attribute
Dim currXMLNode As MSXML.IXMLDOMElement
    Set currXMLNode = oDoc.nodeFromID(mCurrKey)
    currXMLNode.setAttribute AttributeName, NewValue

End Sub
Public Sub ChangeText(NewValue As String)
    Set currXMLNode = oDoc.nodeFromID(mCurrKey)
    currXMLNode.Text = NewValue
End Sub


Public Sub SetTreeView(ByRef oTreeView As TreeView)
'sets an internal reference to the treeview object of the form
    Set mTreeView = oTreeView
End Sub

Public Property Get CurrFileName() As String
'allows to catch the current file name
    CurrFileName = mCurrFileName
End Property
Public Property Get XMLIsLoaded() As Boolean
'is there an XML loaded in the object
    XMLIsLoaded = mXMLIsLoaded
End Property

Public Property Let CurrKey(ByVal vData As String)
'when the user clicks on a node, the instance has to be updated via this property
    mCurrKey = vData
End Property


Public Property Get CurrKey() As String
' to fetch the current key
    CurrKey = mCurrKey
End Property



Public Sub DeleteNode()
'delete a node in the XML tree
Dim currXMLNode As MSXML.IXMLDOMElement
    
    Set currXMLNode = oDoc.nodeFromID(mCurrKey)
    Set nuts = currXMLNode.parentNode.removeChild(currXMLNode)
    mTreeView.Nodes.Remove (mCurrKey)
End Sub

Public Function SaveAsXMLFile(ByVal FilePath As String) As Boolean
'save as function
oDoc.save FilePath
AddLog "XML saved to " & FilePath
End Function

Public Function SaveXMLFile() As Boolean
' save the XML tree to ... an XML file
oDoc.save mCurrFileName
AddLog "XML saved"
End Function

Public Function CloseXMLFile() As Boolean
'set the internal content to nothing
Class_Initialize
AddLog "XML file closed"
End Function

Public Function OpenXMLFile(ByVal FilePath As String) As Boolean
'open an existing xml file and start filling the tree on the form
    Set oDoc = New DOMDocument
    oDoc.async = False
    'oDoc.validateOnParse = False 'If validation is not important, skip it
    oDoc.Load FilePath
    If oDoc.parseError.errorCode = 0 Then
        AddLog "Parsing OK"
         mTreeView.Nodes.Clear
         'this starts the filling process
        AddTreeNode mTreeView, oDoc.documentElement
        AddLog "Tree is loaded in XMLTree instance"
        mCurrFileName = FilePath
        AddLog "Internal file name:" & mCurrFileName
        mXMLIsLoaded = True
        
        OpenXMLFile = True
    Else
        AddLog "Parsing was not successful"
        AddLog oDoc.parseError.reason & vbCrLf & oDoc.parseError.Line & vbCrLf & oDoc.parseError.srcText
        mXMLIsLoaded = False
        OpenXMLFile = False
    End If
End Function
Private Sub AddTreeNode(ByRef objTreeView As TreeView, ByRef oElem As IXMLDOMNode, Optional ByRef oTreeNode As Node)
'this will fill up this object and the treeview on the form as well
    Dim oNewNode As Node
    Dim oNodeList As IXMLDOMNodeList
    Dim i As Long
    'Debug.Print oTreeNode.Text
    If oTreeNode Is Nothing Then
        Set oNewNode = objTreeView.Nodes.Add
    Else
        Set oNewNode = objTreeView.Nodes.Add(oTreeNode, tvwChild)
    End If
    
    oNewNode.Expanded = False
    nodeName = oElem.nodeName
    Select Case nodeName
        Case "folder", "folders"
            use = oElem.Attributes.getNamedItem("DIRNAME").Text
            oNewNode.Image = 1
        Case "file"
            use = oElem.Attributes.getNamedItem("FILENAME").Text
            oNewNode.Image = 3
        Case "TITLE", "URL"
            use = oElem.Text
             oNewNode.Image = 4
    Case Else
        use = "unhandled: " & oElem.nodeName
    End Select
    oNewNode.Text = use
    oNewNode.Key = oElem.Attributes.getNamedItem("ID").Text
    UpdateMaxID oNewNode.Key
    
    Set oNodeList = oElem.childNodes
    If oElem.nodeName = "file" Then
        Set t = objTreeView.Nodes.Add(oNewNode, tvwChild)
        t.Text = oElem.childNodes(0).Text
        t.Key = oElem.childNodes(0).Attributes.getNamedItem("ID").Text
        t.Image = 4
        UpdateMaxID t.Key
        Set t = objTreeView.Nodes.Add(oNewNode, tvwChild)
        t.Text = oElem.childNodes(1).Text
        t.Key = oElem.childNodes(1).Attributes.getNamedItem("ID").Text
        t.Image = 5
        UpdateMaxID t.Key
    Else
        
        For i = 0 To oNodeList.length - 1
            AddTreeNode mTreeView, oNodeList.Item(i), oNewNode
        Next
    End If
End Sub


Public Sub CreateXMLNode(ByVal eNodeType As NodeType, ByVal Name As String, Optional ByVal Title As String, Optional ByVal URL As String)
'this method creates a new file/folder node
Dim currXMLNode As MSXML.IXMLDOMElement
Dim URLNode As MSComctlLib.Node
Dim newXMLnode As MSXML.IXMLDOMElement
Dim IDnode As MSXML.IXMLDOMNode
On Error Resume Next
If mCurrKey = "ID1" Then
    Set currXMLNode = oDoc.documentElement
Else
    Set currXMLNode = oDoc.nodeFromID(mCurrKey)
End If
    Set currTreeNode = mTreeView.Nodes.Item(mCurrKey)
    Set newnode = mTreeView.Nodes.Add(currTreeNode, tvwChild)
    Select Case eNodeType

    Case eFolder 'folder
        newnode.Text = Name
        newnode.Image = 1
        mMaxID = mMaxID + 1
        newnode.Key = "ID" & mMaxID
    
        Set newXMLnode = currXMLNode.appendChild(oDoc.createElement("folder"))
        newXMLnode.setAttribute "ID", "ID" & mMaxID
        newXMLnode.setAttribute "DIRNAME", Name
        
        
    Case eFile 'file
        newnode.Text = Name
        newnode.Image = 3
        mMaxID = mMaxID + 1
        newnode.Key = "ID" & mMaxID
        
        Set titlenode = mTreeView.Nodes.Add(newnode, tvwChild)
        titlenode.Text = Title
        titlenode.Image = 4
        
        Set URLNode = mTreeView.Nodes.Add(newnode, tvwChild)
        URLNode.Text = URL
        URLNode.Image = 5
        
        
        mMaxID = mMaxID + 1
        Set newxmlfilenode = currXMLNode.appendChild(oDoc.createElement("file"))
        newxmlfilenode.setAttribute "ID", "ID" & mMaxID
        titlenode.Key = "ID" & mMaxID
        newxmlfilenode.setAttribute "FILENAME", Name
        
        mMaxID = mMaxID + 1
        Set newxmlTitlenode = newxmlfilenode.appendChild(oDoc.createElement("TITLE"))
        newxmlTitlenode.setAttribute "ID", "ID" & mMaxID
        URLNode.Key = "ID" & mMaxID
        newxmlTitlenode.Text = Title
        
        mMaxID = mMaxID + 1
        Set newxmlURLnode = newxmlfilenode.appendChild(oDoc.createElement("URL"))
        newxmlURLnode.setAttribute "ID", "ID" & mMaxID
        newxmlURLnode.Text = URL
        
        Set currXMLNode = newxmlfilenode
        
    End Select

End Sub
Private Sub UpdateMaxID(ByVal sID As String)
'the maximal ID is not a strict incrementing number but changes when you delete or add nodes
inID = CInt(Right(sID, Len(sID) - 2))
    If inID > mMaxID Then
        mMaxID = inID
    End If
End Sub


Private Sub Class_Initialize()
'standard class ini
    Set oDoc = Nothing
    mCurrKey = 0
    mXMLIsLoaded = False
    mCurrFileName = ""
    mcurrFilepath = ""
    mMaxID = 1
End Sub
