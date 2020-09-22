VERSION 5.00
Begin VB.Form frmQuestions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Questions"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
   Icon            =   "frmQuestions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   480
      Top             =   3360
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "&PREV"
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
      Left            =   600
      Picture         =   "frmQuestions.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
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
      Left            =   2160
      Picture         =   "frmQuestions.frx":3EAC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2895
      Left            =   1800
      TabIndex        =   9
      Top             =   4320
      Width           =   7575
      Begin VB.OptionButton optSel 
         BackColor       =   &H00008080&
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   5
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton optSel 
         BackColor       =   &H00008080&
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton optSel 
         BackColor       =   &H00008080&
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton optSel 
         BackColor       =   &H00008080&
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtQues 
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
         Height          =   495
         Index           =   4
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2040
         Width           =   6135
      End
      Begin VB.TextBox txtQues 
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
         Height          =   495
         Index           =   3
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1440
         Width           =   6135
      End
      Begin VB.TextBox txtQues 
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
         Height          =   495
         Index           =   2
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   840
         Width           =   6135
      End
      Begin VB.TextBox txtQues 
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
         Height          =   495
         Index           =   1
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.TextBox txtQues 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Index           =   0
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Width           =   7575
   End
   Begin VB.Label lblQuesNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "page"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   8880
      TabIndex        =   14
      Top             =   7800
      Width           =   585
   End
   Begin VB.Label lblExamType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "examtype"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   4440
      TabIndex        =   7
      Top             =   120
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUESTION:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   8655
      Left            =   0
      Picture         =   "frmQuestions.frx":6E0F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim SQL As String 'SQL statement to be used
Private totalQues As Integer
Private selQues As Integer
Private j As Integer 'counter
Private k As Integer 'counter
Private score As Integer
Private Answer As String
Private UsrAnswer As String

Private Sub cmdNext_Click()
If cmdNext.Caption = "&Next" Then
    If Not k = Max Then
        k = k + 1 ' order of question
        selQues = QuesID(k - 1) ' value of stored question(id)
        If selQues = 0 Then
            'if no ID is stored in QuesID, create new random question
            Call CreateQues ' 0 is the value of ques id
            Call WriteQues ' write in the database
        Else
            'if there is an ID stored in QuesID, Load question from temp
            Call LoadQues
        End If
    Else
        'if max number, end'
        cmdNext.Caption = "&End"
    End If
Else
    frmRating.Show
    Unload Me
    Exit Sub
End If

lblQuesNo.Caption = k & "/" & Max
End Sub

'load previous existing question from temp
Private Sub cmdPrev_Click()
If Not k = 1 Then
    k = k - 1
    selQues = QuesID(k - 1)
   Call LoadQues
End If
lblQuesNo.Caption = k & "/" & Max
If cmdNext.Caption = "&End" Then cmdNext.Caption = "&Next"
End Sub

Private Sub Form_Load()
DBConnect 'connect to database
AssignID

Select Case ExamQues
Case 0: Max = 10
Case 1: Max = 20
Case 2: Max = 50
End Select

Call DelTemp

j = 0
Call CreateQues
Call WriteQues
lblQuesNo.Caption = k & "/" & Max
End Sub


'assigns ID for all questions
Private Sub AssignID()

'assign SQL statement
Select Case ExamType
Case 0
    SQL = "SELECT * FROM Communications"
    lblExamType.Caption = "Communications"
Case 1
    SQL = "SELECT * FROM Electronics"
    lblExamType.Caption = "Electronics"
Case 2
    SQL = "SELECT * FROM Mathematics"
    lblExamType.Caption = "Mathematics"
End Select

Set Rs = New ADODB.Recordset 'recordset
Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
If Rs.RecordCount <> 0 Then
    totalQues = Rs.RecordCount
    Rs.MoveFirst
    i = 1
    Do While Not Rs.EOF 'while not end of record
        Rs!ID = i 'assigns ID
        Rs.Update
        Rs.MoveNext
        i = i + 1 'increment ID no.
    Loop
End If
Rs.Close
Set Rs = Nothing
End Sub

'create randomized questions
Private Sub CreateQues()
'On Error Resume Next

Do Until ok = True
    Randomize 'initialize random variable (rnd)

    selQues = CInt(Rnd * totalQues) 'round off variable

    For i = 0 To Max - 1 'compare if question exists
        If selQues < 1 Or selQues >= totalQues Then
            ok = False
            Exit For
        End If
        If QuesID(i) = selQues Then
            ok = False
            Exit For
        End If
        If i = Max - 1 And Not QuesID(i) = selQues Then 'if end of max
            ok = True
            QuesID(j) = selQues
            j = j + 1
            k = j
            Exit For
        End If
    Next
Loop

End Sub

'get id and display
Private Sub WriteQues()
For i = 0 To 4
    txtQues(i).Text = ""
Next

'On Error Resume Next
Set Rs = New ADODB.Recordset
Select Case ExamType
Case 0
    SQL = "SELECT * FROM Communications WHERE ID = " & selQues & ""
Case 1
    SQL = "SELECT * FROM Electronics WHERE ID = " & selQues & ""
Case 2
    SQL = "SELECT * FROM Mathematics WHERE ID = " & selQues & ""
End Select
Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
If Rs.RecordCount <> 0 Then
   Answer = Rs!Answer
   txtQues(0).Text = Rs(1)
   For l = 1 To 4
    ok = False
    
    Do Until ok = True
        Randomize
        rndch = CInt(Rnd * 3)
        For i = 1 To 4
            If txtQues(i).Text = Rs(rndch + 2) Then
                ok = False
                Exit For
            End If
            If i = 4 And Not txtQues(i).Text = Rs(rndch + 2) Then
                ok = True
                txtQues(l).Text = Rs(rndch + 2)
                Exit For
            End If
        Next
    Loop
    
  Next
End If
Rs.Close


'store to temp
SQL = "SELECT * FROM Temp"
Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
Rs.AddNew
Rs(0) = selQues
For i = 1 To 5
    Rs(i) = txtQues(i - 1).Text
Next

Rs!Answer = Answer
Rs.Update
Rs.Close
Set Rs = Nothing

Call unclickAll
'optSel(0).Value = True
'Call optSel_Click(0)
End Sub

'load existing question in temp table. the id is to be search in temp table
Private Sub LoadQues()
On Error Resume Next
Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM Temp WHERE ID = " & selQues & ""
Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
If Rs.RecordCount <> 0 Then ' if the recordcount not equal to 0
For i = 1 To 5 ' i represents the data base
    txtQues(i - 1).Text = Rs(i) 'value of the textbox compare to r
Next
Answer = Rs!Answer 'Value of field
optSel(Rs!optSel).Value = True
End If
Rs.Close
Set Rs = Nothing

End Sub

'delete all records in temp table
Private Sub DelTemp()
Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM Temp" ' select all fields from temp table
Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
If Rs.RecordCount <> 0 Then
    Rs.MoveFirst ' go to first record
    Do While Not Rs.EOF 'if not end of file
        Rs.Delete
        Rs.Update
        Rs.MoveNext
    Loop
End If
Rs.Close
Set Rs = Nothing
End Sub

Private Sub optSel_Click(Index As Integer)
UsrAnswer = Trim(txtQues(Index + 1).Text) 'reference for the array
    Set Rs = New ADODB.Recordset
    SQL = "SELECT * FROM Temp WHERE ID = " & selQues & ""
    Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
    If Rs.RecordCount <> 0 Then
        If Trim(UsrAnswer) = Trim(Answer) Then
            Rs!RW = 1
        Else
            Rs!RW = 0
        End If
        Rs!optSel = Index
        Rs.Update
    End If
    Rs.Close
    Set Rs = Nothing
End Sub

Private Sub Timer1_Timer()
For i = 0 To 3
    If optSel(i).Value = True Then
        cmdNext.Enabled = True
        Exit Sub
    End If
Next
cmdNext.Enabled = False
End Sub

Private Sub txtQues_Click(Index As Integer)
If Not Index = 0 Then
optSel(Index - 1).Value = True
optSel(Index - 1).SetFocus
End If
End Sub

Private Sub unclickAll()
For i = 0 To 3
    optSel(i).Value = False
Next
End Sub
