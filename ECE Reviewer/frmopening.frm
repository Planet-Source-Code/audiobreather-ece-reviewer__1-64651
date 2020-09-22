VERSION 5.00
Begin VB.Form frmOpening 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                VIRTUAL REVIEWER                          "
   ClientHeight    =   3165
   ClientLeft      =   3645
   ClientTop       =   3450
   ClientWidth     =   8415
   FontTransparent =   0   'False
   Icon            =   "frmopening.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H000080FF&
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
      Left            =   5880
      Picture         =   "frmopening.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdlogin 
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
      Left            =   3240
      MaskColor       =   &H000000FF&
      Picture         =   "frmopening.frx":3EB6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdregister 
      BackColor       =   &H000080FF&
      Caption         =   "&REGISTER"
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
      MaskColor       =   &H00404040&
      Picture         =   "frmopening.frx":6EA2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   0
      Picture         =   "frmopening.frx":9E8E
      ScaleHeight     =   3195
      ScaleWidth      =   9435
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmOpening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
Unload Me
frmlogo.Show

End Sub

Private Sub cmdLogin_Click()
Unload Me
frmlogin.Show

End Sub

Private Sub cmdregister_Click()
Unload Me
frmreg.Show
End Sub

