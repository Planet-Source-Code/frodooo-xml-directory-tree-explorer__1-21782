VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmViewTree 
   Caption         =   "XML Tree "
   ClientHeight    =   10890
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9495
   Icon            =   "ViewTree.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10890
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtXSL 
      Height          =   285
      Left            =   6780
      TabIndex        =   16
      Text            =   "c:\temp\TreeDesign.xsl"
      Top             =   3120
      Width           =   2625
   End
   Begin VB.CommandButton cmdClearLog 
      Caption         =   "Clear Log"
      Height          =   345
      Left            =   6930
      TabIndex        =   14
      Top             =   8760
      Width           =   1665
   End
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   2025
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   8820
      Width           =   6675
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   7500
      Top             =   7860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "xml"
      DialogTitle     =   "Open XML file"
      Filter          =   "xml"
   End
   Begin VB.Frame frameAdd 
      Caption         =   "Add node"
      Height          =   2715
      Left            =   6750
      TabIndex        =   8
      Top             =   30
      Width           =   2685
      Begin VB.CommandButton cmdAddNode 
         Caption         =   "Add Node"
         Height          =   345
         Left            =   660
         TabIndex        =   7
         Top             =   2250
         Width           =   1905
      End
      Begin VB.OptionButton chkAdd 
         Caption         =   "Add folder"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton chkAdd 
         Caption         =   "add file"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   870
         Width           =   1365
      End
      Begin VB.TextBox txtFolderName 
         Height          =   285
         Left            =   90
         TabIndex        =   2
         Top             =   570
         Width           =   2370
      End
      Begin VB.TextBox txtFileName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   4
         Top             =   1170
         Width           =   1620
      End
      Begin VB.TextBox txtTitle 
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   5
         Top             =   1530
         Width           =   1590
      End
      Begin VB.TextBox txtURL 
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   6
         Top             =   1860
         Width           =   1605
      End
      Begin VB.Label Label2 
         Caption         =   "File name:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "Title:"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1530
         Width           =   1065
      End
      Begin VB.Label Label4 
         Caption         =   "URL:"
         Height          =   405
         Left            =   120
         TabIndex        =   9
         Top             =   1860
         Width           =   1065
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   7860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewTree.frx":0442
            Key             =   "FClosed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewTree.frx":0894
            Key             =   "FOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewTree.frx":3046
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewTree.frx":3498
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewTree.frx":37B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   14843
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label Label5 
      Caption         =   "Path to XSL file:"
      Height          =   255
      Left            =   6750
      TabIndex        =   15
      Top             =   2850
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Log:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   30
      TabIndex        =   13
      Top             =   8460
      Width           =   3975
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuCreateFolder 
         Caption         =   "Create Folder"
      End
      Begin VB.Menu mnuCreateFile 
         Caption         =   "Create File"
      End
      Begin VB.Menu mnuDeleteNode 
         Caption         =   "Delete Node"
      End
      Begin VB.Menu mnuNewXML 
         Caption         =   "New XML"
      End
      Begin VB.Menu mnuLoadXML 
         Caption         =   "Load XML"
      End
      Begin VB.Menu mnuSaveXML 
         Caption         =   "Save XML"
      End
      Begin VB.Menu mnuSaveAsXML 
         Caption         =   "Save XML As"
      End
      Begin VB.Menu mnuCloseXML 
         Caption         =   "Close XML"
      End
   End
End
Attribute VB_Name = "frmViewTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim currTreeNode As MSComctlLib.Node
Private Sub cmdAddNode_Click()
'adds a file or folder to the tree
    If chkAdd(0).Value = True Then
        objXMLTree.CreateXMLNode eFolder, txtFolderName.Text
    Else
        objXMLTree.CreateXMLNode eFile, txtFileName.Text, txtTitle.Text, txtURL.Text
    End If
    txtFolderName.Text = ""
    txtFileName.Text = ""
    txtTitle.Text = ""
    txtURL.Text = ""
End Sub


Private Sub chkAdd_Click(Index As Integer)
'some interface consistency
Select Case Index
Case 0
    txtFolderName.Enabled = True
    
    txtFileName.Enabled = False
    txtTitle.Enabled = False
    txtURL.Enabled = False
    
Case 1
    txtFolderName.Enabled = False
    txtFileName.Enabled = True
    txtTitle.Enabled = True
    txtURL.Enabled = True
    
End Select

End Sub


Private Sub SaveXMLFile()
    objXMLTree.SaveXMLFile
End Sub



Private Sub cmdClearLog_Click()
    txtLog.Text = ""
End Sub

Private Sub Form_Load()
    mnuPopUp.Visible = False
    Set objXMLTree = New XMLTree
    objXMLTree.SetTreeView TreeView
    frmViewTree.Caption = constFormCaption
End Sub

Private Sub mnuCloseXML_Click()
'closes the XML tree
    If objXMLTree.XMLIsLoaded Then
        Response = MsgBox("Save the XML file before closing?", vbYesNo, "CreateXMLTree") 'does the user want to save first
        If Response = vbYes Then
            objXMLTree.SaveXMLFile
        End If
        'clean the interface
        TreeView.Nodes.Clear
        frmViewTree.frameAdd.Enabled = False
        DisableFolderCreation
        DisableFileCreation
        DisableDeleteNode
        DisableCloseXML
        ' close the XML instance
        objXMLTree.CloseXMLFile
        frmViewTree.Caption = constFormCaption
    Else
        MsgBox "No XML loaded yet"
    End If
End Sub

Private Sub mnuCreateFile_Click()
'redirect to the frame
    chkAdd(1).Value = 1
    txtFileName.SetFocus
End Sub

Private Sub mnuCreateFolder_Click()
'redirect to the frame
    chkAdd(0).Value = 1
    txtFolderName.SetFocus
End Sub

Private Sub mnuDeleteNode_Click()
    objXMLTree.DeleteNode
End Sub

Private Sub mnuLoadXML_Click()
'user can open an existing XML file (tree)
    If objXMLTree.XMLIsLoaded Then
        Response = MsgBox("Save the XML file before loading a new one?", vbYesNo, "CreateXMLTree")
        If Response = vbYes Then
            objXMLTree.SaveXMLFile
        End If
    End If
    objXMLTree.CloseXMLFile
    
    ' Set CancelError is True
    Dialog1.CancelError = False
    Dialog1.DialogTitle = "Open XML file"
    
    ' Set flags
    Dialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    Dialog1.Filter = "XML files (*.xml)|*.xml"
    Dialog1.FilterIndex = 0
    
    ' Display the Open dialog box
    Dialog1.ShowOpen
    
    objXMLTree.OpenXMLFile "file:///" & Dialog1.FileName
    
    frmViewTree.Caption = constFormCaption & " (" & Right(objXMLTree.CurrFileName, Len(objXMLTree.CurrFileName) - InStrRev(objXMLTree.CurrFileName, "\")) & ")"
    
End Sub

Private Sub mnuNewXML_Click()
' a brand new tree is created
    If objXMLTree.XMLIsLoaded Then
        Response = MsgBox("Save the XML file before loading a new one?", vbYesNo, "CreateXMLTree")
        If Response = vbYes Then
            objXMLTree.SaveXMLFile
        End If
    End If
    'open a dialog with user
    Dialog1.CancelError = False
    Dialog1.DialogTitle = "New XML file"
    ' Set flags
    Dialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    Dialog1.Filter = "XML files (*.xml)|*.xml"
    Dialog1.FilterIndex = 0
    ' Display the Open dialog box
    Dialog1.ShowOpen
    sFileName = Dialog1.FileName
    'write mendatory header info
    Open sFileName For Output As #1
    Print #1, "<?xml version=""1.0""?>" & vbCrLf & sDTD & vbCrLf
    Print #1, "<?xml-stylesheet type=""text/xsl"" href=""" & txtXSL.Text & """?>" & vbCrLf
    Print #1, "<folders ID=""ID1"" DIRNAME=""New Root"">" & vbCrLf
    Print #1, "<folder ID=""ID2"" DIRNAME=""New Folder"">" & vbCrLf
    Print #1, "<file ID=""ID3"" FILENAME=""New File"">" & vbCrLf
    Print #1, "<TITLE ID=""ID4"">Title</TITLE>" & vbCrLf
    Print #1, "<URL ID=""ID5"">URL</URL>" & vbCrLf
    Print #1, "</file>" & vbCrLf
    Print #1, "</folder>" & vbCrLf
    Print #1, "</folders>" & vbCrLf
    Close #1
     
    objXMLTree.OpenXMLFile sFileName
     
End Sub

Private Sub mnuSaveAsXML_Click()
     Dialog1.CancelError = False
  Dialog1.DialogTitle = "Save XML file"
  ' Set flags
  Dialog1.Flags = cdlOFNHideReadOnly
  ' Set filters
  Dialog1.Filter = "XML files (*.xml)|*.xml"
  Dialog1.FilterIndex = 0
  ' Display the Open dialog box
  Dialog1.ShowSave
objXMLTree.SaveAsXMLFile Dialog1.FileName

End Sub

Private Sub mnuSaveXML_Click()
    objXMLTree.SaveXMLFile
    
End Sub

Private Sub TreeView_AfterLabelEdit(Cancel As Integer, NewString As String)
'you can edit the labels, they refer to XML attributes

Set currTreeNode = TreeView.Nodes.Item(objXMLTree.CurrKey)
Select Case currTreeNode.Image
    Case 1, 2
        objXMLTree.ChangeAttribute "DIRNAME", NewString
    Case 3
        objXMLTree.ChangeAttribute "FILENAME", NewString
    Case 4, 5
        objXMLTree.ChangeText NewString
     
    
End Select

End Sub

Private Sub TreeView_Collapse(ByVal Node As MSComctlLib.Node)
'switches the open/close folder image
Select Case Node.Image
    Case 1
        Node.Image = 2
    Case 2
        Node.Image = 1
    Case 3
End Select
End Sub

Private Sub TreeView_Expand(ByVal Node As MSComctlLib.Node)
'switches the open/close folder image
Select Case Node.Image
    Case 1
        Node.Image = 2
    Case 2
        Node.Image = 1
    Case 3
End Select
End Sub



Private Sub TreeView_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
'handles the right mouse button
    If Button = 2 Then 'if they right click, 1=left, 2=right
        If objXMLTree.XMLIsLoaded Then
            frmViewTree.mnuCloseXML.Enabled = True
            frmViewTree.mnuNewXML.Enabled = True
            frmViewTree.mnuLoadXML.Enabled = True
            frmViewTree.mnuSaveXML.Enabled = True
            frmViewTree.mnuSaveAsXML.Enabled = True
        Else
            frmViewTree.mnuNewXML.Enabled = True
            frmViewTree.mnuLoadXML.Enabled = True
            frmViewTree.mnuSaveXML.Enabled = False
            frmViewTree.mnuSaveAsXML.Enabled = False
            frmViewTree.mnuCloseXML.Enabled = False
            DisableFolderCreation
            DisableFileCreation
            DisableDeleteNode
        End If
        frmViewTree.PopupMenu mnuPopUp 'show popup menu
    Else 'else if they clicked the left button
        DoEvents
    End If
End Sub

Private Sub TreeView_NodeClick(ByVal Node As MSComctlLib.Node)
'determines what you can do depending on the type of node, updates the current key
    Select Case Node.Image
        Case 1, 2 'folder
            CurrKey = Node.Key
            If CurrKey = "ID1" Then ' you cannot delete the root folder
                frmViewTree.mnuDeleteNode.Enabled = False
            Else
                frmViewTree.mnuDeleteNode.Enabled = True
            End If
            
            EnableFolderCreation
            EnableFileCreation
            EnableDeleteNode

        Case 3 ' document
        
            DisableFolderCreation
            DisableFileCreation
            EnableDeleteNode

        Case 4 ' property
        
            DisableFolderCreation
            DisableFileCreation
            DisableDeleteNode
        
    End Select
    'update the key in the instance
    objXMLTree.CurrKey = Node.Key
    AddLog "currTreeNode: " & Node.Key
End Sub

