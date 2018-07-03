VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTool 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10815
   Icon            =   "frmTool.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Check Board Change Shorts"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin VB.TextBox txtShortCloseNodes 
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Text            =   "CadNode.txt"
         Top             =   720
         Width           =   8655
      End
      Begin VB.TextBox txtBoardShorts 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Text            =   "board_short.txt"
         Top             =   240
         Width           =   8655
      End
      Begin VB.CommandButton cmdGoShort 
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8880
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog d1 
         Left            =   960
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strToolPath As String

Dim strShortsCloseNodesPath As String
Dim strBoardShortPath As String



 

Private Sub cmdGoShort_Click()
On Error Resume Next
MkDir strToolPath & "CadNode\TmpNodes"
 
 
If Dir(strShortsCloseNodesPath) = "" Or strShortsCloseNodesPath = "" Then
   MsgBox "Please check " & strShortsCloseNodesPath & "file!", vbCritical
   Exit Sub
End If
 
cmdGoShort.Enabled = False
Call Read_Close_File
 
Call Read_BoardShort_File
 
Call File_Add
Call File_Del
MsgBox "(" & strToolPath & "CadNode\CloseNode.txt) File Creat Ok!", vbInformation
cmdGoShort.Enabled = True
End Sub



Private Sub Form_Load()
On Error Resume Next
strToolPath = App.Path
If Right(strToolPath, 1) <> "\" Then strToolPath = strToolPath & "\"
MkDir strToolPath & "CadNode"

txtBoardShorts.Text = "Please open board_short.txt file!(DblClick me open board_short.txt file!)"
txtShortCloseNodes.Text = "Please open CadNode.txt file!(DblClick me open CadNode.txt file!)"


 
End Sub

 
Private Sub txtBoardShorts_DblClick()
On Error Resume Next
With Me.d1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
  
    
     .Filter = "board_short.txt|*.txt|*.*|*.*"
     .ShowOpen
     txtBoardShorts.Text = Me.d1.FileName
    strBoardShortPath = Me.d1.FileName
    
End With
End Sub






Private Sub Read_Close_File()
Dim strMy As String
 
  Dim i As Integer
   
   
   On Error Resume Next

    Open strShortsCloseNodesPath For Input As #3
 
    
       Do Until EOF(3)
          Line Input #3, strMy
            strMy = Trim(UCase(strMy))
            If strMy <> "" Then
               If Dir(strToolPath & "CadNode\TmpNodes\" & strMy) = "" Then
                  i = i + 1
              
               End If
               Open strToolPath & "CadNode\TmpNodes\" & strMy For Output As #4
               Close #4
               
            End If
 
        DoEvents
       Loop
       
    Close #3
   
End Sub

Private Sub txtShortCloseNodes_DblClick()
'strShortsCloseNosesPath
On Error Resume Next
With Me.d1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
  
    
     .Filter = "*.txt|*.txt|*.*|*.*"
    .ShowOpen
    txtShortCloseNodes.Text = Me.d1.FileName
    strShortsCloseNodesPath = Me.d1.FileName
End With
End Sub
Private Sub Read_BoardShort_File()
Dim strMy As String
 Dim strTmpNode() As String
 Dim strOkNode As String
 Dim strMyStr As String
  Dim i As Integer
Dim bFindOK As Boolean
   
   On Error Resume Next
  Open strToolPath & "CadNode\board_change_short_node.txt" For Output As #5
    Open strBoardShortPath For Input As #3
      Print #5, "!==============board_change_short_node====Start"
    
       Do Until EOF(3)
          Line Input #3, strMy
            strMy = Trim(UCase(strMy))
            If strMy <> "" Then
               strTmpNode = Split(strMy, "->")
               For t = 0 To UBound(strTmpNode)
                 If Dir(strToolPath & "CadNode\TmpNodes\" & Trim(strTmpNode(t))) <> "" Then
                     i = i + 1
                     
                        
                      bFindOK = True
                      If t <> UBound(strTmpNode) Then
                         'strMyStr = Trim(strTmpNode(UBound(strTmpNode))) & "===============>>>>" & Trim(strTmpNode(t))
                         Kill strToolPath & "CadNode\TmpNodes\" & Trim(strTmpNode(t))
                      End If
                 End If

               Next
              If bFindOK = True Then
                 Print #5, Trim(strTmpNode(UBound(strTmpNode)))
                 Open strToolPath & "CadNode\TmpNodes\" & Trim(strTmpNode(UBound(strTmpNode))) For Output As #8
                 Close #8
                 strMyStr = ""
                 bFindOK = False
              End If
            End If
 
        DoEvents
       Loop
       Print #5, "!==============board_change_short_node====End"
    Close #3
    Close #5
End Sub

Private Sub File_Add()
   On Error Resume Next
 Open strToolPath & "CadNode\Nodes.txt" For Output As #5
MyStr = Dir(strToolPath & "CadNode\TmpNodes\*.*")
 
Do While MyStr <> ""
       Print #5, MyStr
      MyStr = Dir
  DoEvents
Loop
Close #5
End Sub
Private Sub File_Del()
On Error Resume Next
Kill strToolPath & "CadNode\TmpNodes\*.*"
RmDir strToolPath & "CadNode\TmpNodes"
End Sub
