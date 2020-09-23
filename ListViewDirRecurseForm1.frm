VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "DriveToTreeToFileToTree By oigres P: Email: oigres@postmaster.co.uk"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   653
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   3735
      Left            =   3240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read file to tree"
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   4440
      Width           =   1815
   End
   Begin ComctlLib.TreeView TreeView2 
      Height          =   3735
      Left            =   6480
      TabIndex        =   2
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6588
      _Version        =   327682
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   6588
      _Version        =   327682
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Choose a different drive"
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label LblPrompt 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "File treeview"
      Height          =   255
      Left            =   6720
      TabIndex        =   5
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Indented list"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Read directory and put it in the treeview and a listbox/file
'Load the  created file to another treeview
'Directory code adapted from MSDN. LoadTreeViewFromFile function from
'http://www.vb-helper.com/HowTo/trvwfile.zip
'Update No2; improved interface; choose a drive;
'can choose treeitem and list children
'
'By oigres P : Email: oigres@postmaster.co.uk

Option Explicit

Dim fnode As node
Dim FIndent As Integer
Dim FIndex As Integer
Dim StrtPath As String

Private Sub Get_Files(FPath As String)
    Dim file_name As String
    Dim File_Path As String
    Dim File_Read As Integer
    Dim x As Boolean, xTemp As Integer, S$
    Dim I As Integer
    On Error Resume Next
    FIndent = FIndent + 1
    File_Path = FPath & "\"
    file_name = Dir$(File_Path, vbDirectory)
    File_Read = 1
    x = False

    Do While file_name <> ""
        If file_name <> "." And file_name <> ".." Then
            If GetAttr(File_Path & file_name) <> vbDirectory Then

                FIndex = FIndex + 1


            Else
                StrtPath = File_Path & file_name

                Set fnode = TreeView1.Nodes.Add(File_Path, tvwChild, FPath & "\", file_name)

                'changed to dash/hyphen for readability
                Text1.Text = Text1.Text & "->" & String(FIndent * FIndent, "_") & file_name & vbCrLf
                S$ = ""
                ''possible hack ; if FIndent=1 this doesn't execute
                For xTemp = 2 To FIndent
                    S$ = S$ & vbTab
                Next
                'write to a file; number of levels= number of tabs
                Print #1, S$ & file_name

                FIndex = FIndex + 1
                x = True
                'recursive call
                Get_Files StrtPath

            End If

        End If
        If x = True Then
            file_name = Dir$(File_Path, vbDirectory)
            For I = 2 To File_Read
                file_name = Dir$
            Next
            x = False
        End If
        file_name = Dir$
        File_Read = File_Read + 1

    Loop
    FIndent = FIndent - 1

End Sub


Private Sub Command1_Click()
    'read in the file to a tree
    TreeView2.Visible = False
    LoadTreeViewFromFile App.Path & "\" & "Mytest.txt", TreeView2
    TreeView2.Visible = True
End Sub

Private Sub Drive1_Change()
    Dim x As Integer, S$
    Open App.Path & "\" & "Mytest.txt" For Output As #1

    Text1.Text = ""
    TreeView1.Nodes.Clear

    Text1.Text = Text1.Text & Left$(Drive1.Drive, 2) & "\" & vbCrLf
    ' need \ on end of key here------------|
    Set fnode = TreeView1.Nodes.Add(, , Left$(Drive1.Drive, 2) & "\", Left$(Drive1.Drive, 2) & "\")
    'initial data to file; topmost root of tree
    Print #1, Left$(Drive1.Drive, 2) & "\"
    'initialise variables
    FIndent = 1
    FIndex = 0
    Text1.Visible = False
    TreeView1.Visible = False
    LblPrompt.Caption = "Processing drive..."
    LblPrompt.Refresh
    'note does not end with '\'
    StrtPath = Left$(Drive1.Drive, 2)
    'start the directory reading
    Get_Files StrtPath
    Text1.Visible = True
    TreeView1.Visible = True
    LblPrompt.Caption = "Finished"

    Close 1
End Sub

Private Sub Form_Load()

    Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'maynot be necessary
    Close #1
End Sub

Private Sub LoadTreeViewFromFile(ByVal file_name As String, ByVal trv As TreeView)
    Dim fnum As Integer
    Dim text_line As String
    Dim level As Integer
    Dim tree_nodes() As node
    Dim num_nodes As Integer

    fnum = FreeFile
    Open file_name For Input As fnum

    trv.Nodes.Clear
    Do While Not EOF(fnum)
        ' Get a line.
        Line Input #fnum, text_line

        ' Find the level of indentation.
        level = 1
        Do While Left$(text_line, 1) = vbTab
            level = level + 1
            text_line = Mid$(text_line, 2)
        Loop

        ' Make room for the new node.
        If level > num_nodes Then
            num_nodes = level
            ReDim Preserve tree_nodes(1 To num_nodes)
        End If

        ' Add the new node.
        If level = 1 Then
            Set tree_nodes(level) = trv.Nodes.Add(, , , text_line)
        Else
            Set tree_nodes(level) = trv.Nodes.Add(tree_nodes(level - 1), tvwChild, , text_line)
            'tons faster without this
            ''tree_nodes(level).EnsureVisible
        End If
    Loop

    Close fnum

End Sub


Private Sub TreeView1_Click()
    On Error GoTo ehandler
    Dim node As New node
    Set node = TreeView1.SelectedItem
    TreeView1.ToolTipText = node.Key
    Exit Sub
ehandler:
    MsgBox Err.Description & ":" & Err.Number & ":" & Err.LastDllError
End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
    Dim node As New node, fnode As New node

    If KeyAscii = vbKeyReturn Then
        Set node = TreeView1.SelectedItem
        Open App.Path & "\" & "Mytest.txt" For Output As #1

        Text1.Text = ""
        TreeView1.Nodes.Clear

        Text1.Text = Text1.Text & node.Key & vbCrLf
        Set fnode = TreeView1.Nodes.Add(, , node.Key, node.Key)

        'initial data to file; topmost root of tree
        Print #1, node.Key

        FIndent = 1
        FIndex = 0
        Text1.Visible = False
        TreeView1.Visible = False
        LblPrompt.Caption = "Processing drive..."
        LblPrompt.Refresh
        'note does not end with '\';item key does end with \
        StrtPath = Mid$(node.Key, 1, Len(node.Key) - 1)

        'start the directory reading
        Get_Files StrtPath
        Text1.Visible = True
        TreeView1.Visible = True
        LblPrompt.Caption = "Finished"
        Close #1
    End If

End Sub

