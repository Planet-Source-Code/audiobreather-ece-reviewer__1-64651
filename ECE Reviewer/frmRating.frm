VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRating 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rating"
   ClientHeight    =   10665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15240
   Icon            =   "frmRating.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10665
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      Caption         =   "&END REVIEW"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      Picture         =   "frmRating.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9600
      Width           =   2415
   End
   Begin VB.CommandButton cmdanother 
      Caption         =   "&TAKE ANOTHER EXAM"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      Picture         =   "frmRating.frx":3EB6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9600
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7575
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   13361
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   1095
      Left            =   13200
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.Label lblPercent 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   675
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1440
      End
   End
   Begin VB.Label lblCor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   12120
      TabIndex        =   1
      Top             =   360
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Correct Answers:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   9240
      TabIndex        =   8
      Top             =   350
      Width           =   2775
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ExamQues"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Questions"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5400
      TabIndex        =   4
      Top             =   345
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   10635
      Left            =   0
      Picture         =   "frmRating.frx":6EA2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rating:"
      Height          =   195
      Left            =   8040
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmRating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim SQL As String



Private Sub cmdanother_Click()
    For i = 0 To 49
        QuesID(i) = 0
    Next
Load frmmain
frmmain.Show
Unload Me
End Sub


Private Sub Form_Load()

With ListView1
    .ListItems.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Questions", 10700
    .ColumnHeaders.Add , , "Your Answer", 2100
    .ColumnHeaders.Add , , "Correct Answer", 2100

End With

Select Case ExamType
Case 0: lbl2.Caption = "Communications"
Case 1: lbl2.Caption = "Electronics"
Case 2: lbl2.Caption = "Mathematics"
End Select

Select Case ExamQues
Case 0: lbl1.Caption = "10 questions"
Case 1: lbl1.Caption = "20 questions"
Case 2: lbl1.Caption = "50 questions"
End Select

DBConnect
Set Rs = New ADODB.Recordset
SQL = "SELECT SUM(RW) as TOTAL FROM Temp"
Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
If Rs.RecordCount <> 0 Then
    lblPercent.Caption = CInt((Rs!TOTAL / Max) * 100) & "%"
    lblCor.Caption = Rs!TOTAL & "/" & Max
End If
Rs.Close

j = 1 ' counter for the listview
ListView1.ListItems.Clear
For i = 0 To Max - 1
    SQL = "SELECT * FROM Temp WHERE ID = " & QuesID(i) & ""
    Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
    If Rs.RecordCount <> 0 Then
        Rs.MoveFirst
        With ListView1
            .ListItems.Add , , Rs!Question
            .ListItems(j).ListSubItems.Add , , Rs(Rs!optSel + 2)
            .ListItems(j).ListSubItems.Add , , Rs!Answer
        End With
    End If
    Rs.Close
    j = j + 1
Next
Set Rs = Nothing

Call SaveToHistory

End Sub

Private Sub Form_Resize()
ListView1.Width = Me.Width * 0.98
End Sub

Private Sub cmdexit_Click()
Dim resp As Integer
resp = MsgBox("Do you really want to END this review?", vbYesNo + vbQuestion, "----------REVIEWER----------")
If resp = vbYes Then
Unload Me
End If
End Sub

Private Sub SaveToHistory()
On Error Resume Next
Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM History"
Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
Rs.AddNew
Rs(0) = Username

Select Case ExamType
Case 0: Rs(1) = "Communications"
Case 1: Rs(1) = "Electronics"
Case 2: Rs(1) = "Mathematics"
End Select

Rs(2) = Max
Rs(3) = lblPercent.Caption
Rs(4) = Format(Date, "dd/mm/yyyy") & ", " & Time
Rs.Update
Rs.Close
Set Rs = Nothing
End Sub

