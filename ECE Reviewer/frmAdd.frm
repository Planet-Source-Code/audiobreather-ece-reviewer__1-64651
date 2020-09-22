VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Questions"
   ClientHeight    =   9150
   ClientLeft      =   3225
   ClientTop       =   2295
   ClientWidth     =   10320
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtQues 
      BackColor       =   &H00FFFFC0&
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
      Left            =   3360
      TabIndex        =   15
      Top             =   6840
      Width           =   6135
   End
   Begin VB.TextBox txtQues 
      BackColor       =   &H00FFFFC0&
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
      Left            =   3360
      TabIndex        =   14
      Top             =   6120
      Width           =   6135
   End
   Begin VB.TextBox txtQues 
      BackColor       =   &H00FFFFC0&
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
      Left            =   3360
      TabIndex        =   13
      Top             =   5400
      Width           =   6135
   End
   Begin VB.TextBox txtQues 
      BackColor       =   &H00FFFFC0&
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
      Left            =   3360
      TabIndex        =   12
      Top             =   4680
      Width           =   6135
   End
   Begin VB.TextBox txtQues 
      BackColor       =   &H00FFFFC0&
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
      Index           =   5
      Left            =   3360
      TabIndex        =   11
      Top             =   7560
      Width           =   6135
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8040
      Top             =   6120
   End
   Begin VB.CommandButton cmdQues 
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
      Index           =   5
      Left            =   360
      Picture         =   "frmAdd.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8400
      Width           =   1575
   End
   Begin VB.ComboBox cmbType 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmAdd.frx":3EAC
      Left            =   6600
      List            =   "frmAdd.frx":3EB9
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdQues 
      Caption         =   "&DELETE"
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
      Index           =   4
      Left            =   240
      Picture         =   "frmAdd.frx":3EE7
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdQues 
      Caption         =   "&EDIT"
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
      Index           =   3
      Left            =   240
      Picture         =   "frmAdd.frx":6D61
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdQues 
      Caption         =   "&NEW"
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
      Index           =   2
      Left            =   240
      Picture         =   "frmAdd.frx":9BDB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdQues 
      Caption         =   "&CANCEL"
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
      Index           =   1
      Left            =   8280
      Picture         =   "frmAdd.frx":CA55
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdQues 
      Caption         =   "&SAVE"
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
      Index           =   0
      Left            =   6720
      Picture         =   "frmAdd.frx":F8CF
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8400
      Width           =   1575
   End
   Begin VB.TextBox txtQues 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3360
      Width           =   7455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choice 1:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   2160
      TabIndex        =   20
      Top             =   4800
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choice 2:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   2160
      TabIndex        =   19
      Top             =   5520
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choice 3:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   2160
      TabIndex        =   18
      Top             =   6240
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choice 4:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   3
      Left            =   2160
      TabIndex        =   17
      Top             =   6960
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Answer:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   270
      Index           =   4
      Left            =   2160
      TabIndex        =   16
      Top             =   7680
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Question:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   1665
   End
   Begin VB.Label lblExamType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "examtype"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   4920
      TabIndex        =   8
      Top             =   3000
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   9120
      Left            =   0
      Picture         =   "frmAdd.frx":12749
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10320
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim SQL As String
Private QuesID As String
Private NewEdit As Integer
Private NoAns As Boolean
Private SameCh As Boolean

Private Sub cmbType_Click()
lblExamType.Caption = cmbType.Text
Call QuesList
End Sub

Private Sub cmdQues_Click(Index As Integer)
Select Case Index
Case 0: 'save question
    For i = 1 To 4
        If Trim(txtQues(5).Text) = Trim(txtQues(i).Text) Then
            NoAns = False
            Exit For
        Else
            NoAns = True
        End If
    Next
        
    If NoAns = True Then
        MsgBox "Please input the answer based on the choices given!", vbOKOnly + vbExclamation, "Reviewer"
        txtQues(5).SetFocus
        SendKeys "{HOME}+{END}"
        Exit Sub
    End If
    
    For i = 1 To 4
        For j = 1 To 4
            If i = j Then GoTo nextL
            If txtQues(i).Text = txtQues(j).Text Then
                MsgBox "Same Choices are not allowed!", vbOKOnly + vbExclamation, "Reviewer"
                txtQues(i).SetFocus
                SendKeys "{HOME}+{END}"
                Exit Sub
            End If
nextL:
        Next
    Next
        
    If NewEdit = 0 Then 'new question
        Set Rs = New ADODB.Recordset
        SQL = "SELECT * FROM " & cmbType.Text & " WHERE Question = '" & txtQues(0).Text & "'"
        Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
        If Rs.RecordCount = 0 Then
            Rs.AddNew
            For i = 1 To 6
                Rs(i) = txtQues(i - 1)
            Next
            Rs.Update
            Rs.Close
            Set Rs = Nothing
            Call QuesList
            MsgBox "Question Added!", vbOKOnly + vbInformation, "Reviewer"
            Exit Sub
        Else
            MsgBox "Question already exists in database!", vbOKOnly + vbExclamation, "Reviewer"
            txtQues(0).SetFocus
            SendKeys "{HOME}+{END}"
        End If
        Rs.Close
        Set Rs = Nothing
    ElseIf NewEdit = 1 Then 'edit question
        Set Rs = New ADODB.Recordset
        SQL = "SELECT * FROM " & cmbType.Text & " WHERE ID = " & QuesID & ""
        Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
        If Rs.RecordCount <> 0 Then
            For i = 1 To 6
                Rs(i) = txtQues(i - 1)
            Next
            Rs.Update
            Rs.Close
            Set Rs = Nothing
            Call QuesList
            MsgBox "Question updated!", vbOKOnly + vbInformation, "Reviewer"
            Exit Sub
        Else
            MsgBox "Question already exists in database!", vbOKOnly + vbExclamation, "Reviewer"
            txtQues(0).SetFocus
            SendKeys "{HOME}+{END}"
        End If
        Rs.Close
        Set Rs = Nothing
    End If
Case 1: 'cancel
    NewEdit = 2
    ListView1.Enabled = True
    cmdQues(0).Enabled = False
    cmdQues(1).Enabled = False
    cmdQues(2).Enabled = True
    cmdQues(3).Enabled = True
    cmdQues(4).Enabled = True
    Call QuesList
    Call DisableFields
    Timer1.Enabled = False
Case 2: 'new
    NewEdit = 0
    Call EnableFields
    ListView1.Enabled = False
    cmdQues(0).Enabled = True
    cmdQues(1).Enabled = True
    cmdQues(2).Enabled = False
    cmdQues(3).Enabled = False
    cmdQues(4).Enabled = False
    Call ClearAll
    txtQues(0).SetFocus
    Timer1.Enabled = True
Case 3: 'edit
    NewEdit = 1
    Call EnableFields
    ListView1.Enabled = False
    cmdQues(0).Enabled = True
    cmdQues(1).Enabled = True
    cmdQues(2).Enabled = False
    cmdQues(3).Enabled = False
    cmdQues(4).Enabled = False
    ListView1.Enabled = False
    txtQues(0).SetFocus
    Timer1.Enabled = True
Case 4: 'delete
    If MsgBox("Are you sure?", vbYesNo + vbQuestion, "Reviewer") = vbYes Then
        Set Rs = New ADODB.Recordset
        SQL = "SELECT * FROM " & cmbType.Text & ""
        Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
        If Rs.RecordCount <> 0 Then
            Rs.Delete
            Rs.Update
        End If
        Rs.Close
        Set Rs = Nothing
        Call QuesList
    End If
Case 5: 'back
    frmlogin.Show
    Unload Me
End Select
End Sub

Private Sub Form_Load()
DBConnect
With ListView1
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "ID", 800
    .ColumnHeaders.Add , , "Questions", 5000
    .ColumnHeaders.Add , , "Answer", 1200
End With
cmbType.ListIndex = 0
Call cmdQues_Click(1)
End Sub

Private Sub ClearAll()
For i = 0 To 5
    txtQues(i).Text = ""
Next
End Sub

Private Sub QuesList()
Call AssignID
ListView1.ListItems.Clear
Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM " & cmbType.Text & ""
Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
If Rs.RecordCount <> 0 Then
    With ListView1
    Rs.MoveFirst
    j = 1
    Do While Not Rs.EOF
        .ListItems.Add , , Rs!ID
        .ListItems(j).ListSubItems.Add , , Rs!Question
        .ListItems(j).ListSubItems.Add , , Rs!Answer
        j = j + 1
        Rs.MoveNext
    Loop
    End With
End If
Rs.Close
Set Rs = Nothing
Call ListView1_Click
End Sub





Private Sub ListView1_Click()
On Error Resume Next
If ListView1.Enabled = True Then
    QuesID = ListView1.SelectedItem.Text
    Call LoadInfo
End If
End Sub

Private Sub LoadInfo()
On Error Resume Next
Call ClearAll

Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM " & cmbType.Text & " WHERE ID = " & ListView1.SelectedItem.Text & ""
Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
If Rs.RecordCount <> 0 Then
    Rs.MoveFirst
    For i = 1 To 6
        txtQues(i - 1).Text = Rs(i)
    Next
End If
Rs.Close
Set Rs = Nothing
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
If ListView1.Enabled = True Then
    QuesID = ListView1.SelectedItem.Text
    Call LoadInfo
End If
End Sub

Private Sub AssignID()

Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM " & cmbType.Text & ""
Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
If Rs.RecordCount <> 0 Then
    totalQues = Rs.RecordCount
    Rs.MoveFirst
    i = 1
    Do While Not Rs.EOF
        Rs!ID = i
        Rs.Update
        Rs.MoveNext
        i = i + 1
    Loop
End If
Rs.Close
Set Rs = Nothing
End Sub

Private Sub DisableFields()
For i = 0 To 5
    txtQues(i).Locked = True
Next
End Sub

Private Sub EnableFields()
For i = 0 To 5
    txtQues(i).Locked = False
Next
End Sub

Private Sub Timer1_Timer()
For i = 0 To 5
    If Trim(txtQues(i).Text) = "" Then
        cmdQues(0).Enabled = False
        Exit Sub
    End If
Next
cmdQues(0).Enabled = True
End Sub


