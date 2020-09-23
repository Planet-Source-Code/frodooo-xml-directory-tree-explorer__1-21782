Attribute VB_Name = "Global"

Public objXMLTree As XMLTree
Public LogCounter As Integer
Public Const sDTD = "<!DOCTYPE folders[ " & vbCrLf & _
    "<!ELEMENT folders (folder|file)+>" & vbCrLf & _
    "<!ELEMENT folder (file| folder)*>" & vbCrLf & _
    "<!ELEMENT file (TITLE, URL)>" & vbCrLf & _
    "<!ELEMENT TITLE (#PCDATA)>" & vbCrLf & _
    "<!ELEMENT URL (#PCDATA)>" & vbCrLf & _
    "<!ATTLIST folders" & vbCrLf & _
    "DIRNAME CDATA #REQUIRED" & vbCrLf & _
    "ID ID #REQUIRED" & vbCrLf & _
    ">" & vbCrLf & _
    "<!ATTLIST folder" & vbCrLf & _
    "DIRNAME CDATA #REQUIRED" & vbCrLf & _
    "ID ID #REQUIRED" & vbCrLf & _
    ">" & vbCrLf & _
    "<!ATTLIST file" & vbCrLf & _
    "FILENAME CDATA #REQUIRED" & vbCrLf & _
    "ID ID #REQUIRED" & vbCrLf & _
">" & vbCrLf & _
"<!ATTLIST TITLE" & vbCrLf & _
"ID ID #REQUIRED" & vbCrLf & _
">" & vbCrLf & _
"<!ATTLIST URL" & vbCrLf & _
"ID ID #REQUIRED" & vbCrLf & _
">" & vbCrLf & _
"]>"
Public Const constFormCaption = "XML Tree"
Public Enum NodeType
    eFolder = 0
    eFile = 1
    eAttribute = 2
End Enum

Public Sub EnableCloseXML()
    frmViewTree.mnuCloseXML.Enabled = True
End Sub
Public Sub DisableCloseXML()
    frmViewTree.mnuCloseXML.Enabled = False
End Sub


Public Sub EnableFolderCreation()
    frmViewTree.frameAdd.Enabled = True
    frmViewTree.mnuCreateFolder.Enabled = True
End Sub
Public Sub DisableFolderCreation()
    frmViewTree.frameAdd.Enabled = False
    frmViewTree.mnuCreateFolder.Enabled = False
End Sub
Public Sub EnableFileCreation()
    frmViewTree.mnuCreateFile.Enabled = True
End Sub
Public Sub DisableFileCreation()
    frmViewTree.mnuCreateFile.Enabled = False
End Sub
Public Sub EnableDeleteNode()
    frmViewTree.mnuDeleteNode.Enabled = True
End Sub
Public Sub DisableDeleteNode()
    frmViewTree.mnuDeleteNode.Enabled = False
End Sub

Public Sub AddLog(s As String)
    LogCounter = LogCounter + 1
    frmViewTree.txtLog.Text = frmViewTree.txtLog.Text & LogCounter & "> " & s & vbCrLf
End Sub
