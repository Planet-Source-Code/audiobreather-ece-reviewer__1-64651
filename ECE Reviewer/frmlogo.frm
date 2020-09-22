VERSION 5.00
Begin VB.Form frmLogo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WELCOME"
   ClientHeight    =   8565
   ClientLeft      =   3645
   ClientTop       =   3450
   ClientWidth     =   8985
   Icon            =   "frmlogo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Picture         =   "frmlogo.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton cmdenter 
      Caption         =   "&ENTER"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Picture         =   "frmlogo.frx":3EB6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   8565
      Left            =   0
      Picture         =   "frmlogo.frx":6EA2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmlogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdenter_Click()
Unload Me
frmOpening.Show

End Sub

Private Sub cmdexit_Click()

Dim resp As Integer
resp = MsgBox("Do you really want to exit?", vbYesNo + vbQuestion, "----------REVIEWER----------")
If resp = vbYes Then
Unload Me

End If
End Sub


