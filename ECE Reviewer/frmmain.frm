VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmmain.frx":0ECA
   ScaleHeight     =   7365
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHistory 
      Caption         =   "&History"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Picture         =   "frmmain.frx":F6BF
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
      Caption         =   "&BACK"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Picture         =   "frmmain.frx":12539
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   720
      TabIndex        =   6
      Top             =   2280
      Width           =   6375
      Begin VB.OptionButton opt1 
         BackColor       =   &H00008080&
         Caption         =   "Mathematics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   4680
         TabIndex        =   9
         Top             =   120
         Width           =   1695
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00008080&
         Caption         =   "Electronics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   8
         Top             =   120
         Width           =   1815
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00008080&
         Caption         =   "Communications"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   3840
      Width           =   6375
      Begin VB.OptionButton opt2 
         BackColor       =   &H00008080&
         Caption         =   "10 Questions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   2055
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00008080&
         Caption         =   "20 Questions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   4
         Top             =   120
         Width           =   1815
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00008080&
         Caption         =   "50 Questions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   4680
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      Picture         =   "frmmain.frx":1551B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label lblselect 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Exam"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   2640
      TabIndex        =   0
      Top             =   1080
      Width           =   2385
   End
   Begin VB.Image Image1 
      Height          =   7335
      Left            =   0
      Picture         =   "frmmain.frx":18395
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExam_Click(Index As Integer)

End Sub

Private Sub cmdBack_Click()
frmOpening.Show
Unload Me
End Sub

Private Sub cmdHistory_Click()
frmHistory.Show vbModal
End Sub

Private Sub cmdOk_Click()
Load frmQuestions
frmQuestions.Show
Unload Me
End Sub

Private Sub Form_Load()
DBConnect
opt1(0).Value = True
opt2(0).Value = True
opt1_Click (0)
opt2_Click (0)
End Sub


Private Sub opt1_Click(Index As Integer)
Select Case Index
Case 0:
   ExamType = 0
Case 1:
   ExamType = 1
Case 2:
   ExamType = 2
End Select
End Sub

Private Sub opt2_Click(Index As Integer)
Select Case Index
Case 0:
   ExamQues = 0
Case 1:
   ExamQues = 1
Case 2:
   ExamQues = 2
End Select
End Sub
