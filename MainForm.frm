VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Classic ToDo"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Control ToDo"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   3735
      Begin VB.CommandButton Command3 
         Caption         =   "Complete"
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete ToDo"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ToDo Lists"
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   3735
      Begin VB.ListBox List1 
         Height          =   2790
         ItemData        =   "MainForm.frx":0000
         Left            =   120
         List            =   "MainForm.frx":0002
         TabIndex        =   4
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add ToDo"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton Command1 
         Caption         =   "Add ToDo"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call AddToDo
End Sub

Private Sub AddToDo()
    If Text1.Text = "" Then
        MsgBox "Type ToDo First!"
    Else
        List1.AddItem (Text1.Text)
        Text1.Text = ""
        Text1.SetFocus
    End If
End Sub

Private Sub Command2_Click()
    'MsgBox List1.ListIndex
    
    If List1.ListIndex = -1 Then
        MsgBox "Select ToDo First!"
    Else
        List1.RemoveItem (List1.ListIndex)
    End If
End Sub

Private Sub Command3_Click()
    Dim search As Integer
    search = InStr(1, List1.Text, "--")
    
    If List1.ListIndex = -1 Then
        MsgBox "Select ToDo First!"
    ElseIf search > 0 Then
        MsgBox "The ToDo has been completed before"
    Else
        List1.List(List1.ListIndex) = "--" + List1.Text + "-- (Completed)"
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call AddToDo
        KeyAscii = 0
    End If
End Sub
