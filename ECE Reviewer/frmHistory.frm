VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "History"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   Icon            =   "frmHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Cancel          =   -1  'True
      Caption         =   "&Clear History"
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
      Left            =   5760
      Picture         =   "frmHistory.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Image Image1 
      Height          =   4845
      Left            =   0
      Picture         =   "frmHistory.frx":3EB6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7920
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim SQL As String

Private Sub cmdclear_Click()
Set Rs = New ADODB.Recordset
SQL = "SELECT * FROM History WHERE Username = '" & Username & "'"
Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
If Rs.RecordCount > 0 Then
    Rs.MoveFirst
    Do While Not Rs.EOF
        Rs.Delete
        Rs.MoveNext
    Loop
End If
Rs.Close
Set Rs = Nothing
Call LoadHistory
MsgBox "History cleared!", vbOKOnly + vbInformation, "ECE Reviewer"
End Sub

Private Sub Form_Load()
DBConnect
Me.Caption = "Exam History of " & Username & ""
With ListView1
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Category", 0.25 * ListView1.Width
    .ColumnHeaders.Add , , "Num of items", 0.23 * ListView1.Width
    .ColumnHeaders.Add , , "Percentage", 0.2 * ListView1.Width
    .ColumnHeaders.Add , , "Date and Time", 0.3 * ListView1.Width
End With
Image1.Height = Me.Height
Image1.Width = Me.Width

Call LoadHistory
End Sub

Private Sub LoadHistory()
Set Rs = New ADODB.Recordset
ListView1.ListItems.Clear
SQL = "SELECT * FROM History WHERE Username = '" & Username & "'"
Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
If Rs.RecordCount > 0 Then
    With ListView1
        j = 1
        Rs.MoveFirst
        Do While Not Rs.EOF
            .ListItems.Add , , Rs!Category
            For i = 2 To 4
                .ListItems(j).ListSubItems.Add , , Rs(i)
            Next
            j = j + 1
            Rs.MoveNext
        Loop
    End With
End If
Rs.Close
Set Rs = Nothing
End Sub

