VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EVR-login"
   ClientHeight    =   3165
   ClientLeft      =   3645
   ClientTop       =   3450
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
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
      Left            =   6000
      Picture         =   "frmlogin.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H000080FF&
      Caption         =   "&LOGIN"
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
      Left            =   600
      Picture         =   "frmlogin.frx":3EB6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3600
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label lblpassword 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblname3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   3240
      Left            =   0
      Picture         =   "frmlogin.frx":6EA2
      Top             =   0
      Width           =   8640
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim SQL As String

Private Sub cmdcancel_Click()
txtName.Text = ""
txtName.SetFocus
txtPassword.Text = ""
frmlogin.Hide
frmOpening.Show
Unload Me
End Sub

Private Sub cmdLogin_Click()
If txtName = "" And txtPassword = "" Then
    MsgBox "please supply the required data", vbOKOnly + vbExclamation, "error"
    txtName.SetFocus
ElseIf txtName = "" Then
    MsgBox "please supply username", vbOKOnly + vbExclamation, "error"
    txtName.SetFocus
ElseIf txtPassword = "" Then
    MsgBox "missing password", vbOKOnly + vbExclamation, "error"
    txtPassword.SetFocus
Else
    Set Rs = New ADODB.Recordset
    SQL = "SELECT * FROM Login WHERE Username = '" & txtName.Text & "'"
    Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
    If Rs.RecordCount <> 0 Then
        Rs.MoveFirst
        If Trim(txtPassword.Text) = Rs!Password Then
            MsgBox "WELCOME TO VIRTUAL REVIEWER", vbOKOnly + vbInformation, "----------------REVIEWER---------------"
            If Trim(txtName.Text) = "admin" Then
                frmAdd.Show
                Unload Me
            Else
                Username = Trim(txtName.Text)
                frmmain.Show
                Unload Me
            End If
        Else
            MsgBox "Wrong password!", vbOKOnly + vbExclamation, "Reviewer"
            txtPassword.Text = ""
            txtPassword.SetFocus
        End If
    Else
        MsgBox "Username not found!", vbOKOnly + vbExclamation, "Reviewer"
        txtName.Text = ""
        txtPassword.Text = ""
        txtName.SetFocus
    End If
    Rs.Close
    Set Rs = Nothing
End If

End Sub

Private Sub Form_Load()
DBConnect
End Sub


