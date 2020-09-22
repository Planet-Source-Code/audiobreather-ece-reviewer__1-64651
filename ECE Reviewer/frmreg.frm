VERSION 5.00
Begin VB.Form frmReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                     REGISTRATION"
   ClientHeight    =   3165
   ClientLeft      =   3645
   ClientTop       =   3450
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmreg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "C&ANCEL"
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
      Left            =   5880
      Picture         =   "frmreg.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&LEAR"
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
      Left            =   3240
      Picture         =   "frmreg.frx":3EB6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "&CONFIRM"
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
      Left            =   480
      Picture         =   "frmreg.frx":6EA2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "Password"
      DataMember      =   "2"
      DataSource      =   "UserInfo"
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
      DataField       =   "Username"
      DataMember      =   "1"
      DataSource      =   "UserInfo"
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
      MaxLength       =   12
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label lblpassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter password"
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
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lblname1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter username"
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
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   3240
      Left            =   0
      Picture         =   "frmreg.frx":9E8E
      Top             =   0
      Width           =   8640
   End
End
Attribute VB_Name = "frmreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim SQL As String
Private Sub Form_Load()
DBConnect
End Sub
Private Sub cmdcancel_Click()
Me.Hide
frmOpening.Show
txtName = ""
txtPassword = ""
End Sub

Private Sub cmdclear_Click()
txtName.Text = ""
txtPassword.Text = ""
txtName.SetFocus

End Sub

Private Sub cmdConfirm_Click()
If txtName.Text = "" Then
MsgBox "please supply username", vbOKOnly + vbExclamation, "error"
txtName.SetFocus
ElseIf txtPassword.Text = "" Then
MsgBox "missing password", vbOKOnly + vbExclamation, "error"
txtPassword.SetFocus
ElseIf txtName.Text = "" And txtPassword.Text = "" Then
MsgBox "please supply the required field", vbOKOnly + vbCancel, "error"
Else
        Set Rs = New ADODB.Recordset
        SQL = "SELECT * FROM Login WHERE Username = '" & Trim(txtName.Text) & "'"
        Rs.Open SQL, dbCN, adOpenKeyset, adLockOptimistic
        If Rs.RecordCount <> 0 Then
            MsgBox "Username already exists in database!", vbOKOnly + vbExclamation, "Reviewer"
            txtName = ""
            txtPassword = ""
            txtName.SetFocus
        Else
            Rs.AddNew
            Rs!UserName = Trim(txtName.Text)
            Rs!Password = Trim(txtPassword.Text)
            Rs.Update
            MsgBox "User added to database!", vbOKOnly + vbInformation, "Reviewer"
            Unload Me
            frmlogin.Show
        End If
        Rs.Close
        Set Rs = Nothing

    End If

    End Sub



