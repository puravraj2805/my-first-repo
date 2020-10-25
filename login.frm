VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9660
   ClientLeft      =   4305
   ClientTop       =   1035
   ClientWidth     =   15240
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   15240
   Begin VB.CommandButton Command7 
      Caption         =   "USER"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9720
      TabIndex        =   8
      Top             =   5520
      Width           =   4000
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ADMIN"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3120
      TabIndex        =   7
      Top             =   5520
      Width           =   4000
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CONTACT US"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   8001
      TabIndex        =   5
      Top             =   3000
      Width           =   2000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ABOUT US"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6001
      TabIndex        =   4
      Top             =   3000
      Width           =   2000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BLOG"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4001
      TabIndex        =   3
      Top             =   3000
      Width           =   2000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TOURS"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2001
      TabIndex        =   2
      Top             =   3000
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   2000
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   0
      Picture         =   "login.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   16035
      TabIndex        =   0
      Top             =   0
      Width           =   16095
   End
   Begin VB.Label Label1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   6
      Top             =   4080
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command6_Click()
Unload Form2
Form2.Show
End Sub

Private Sub Command7_Click()
Unload Form3
Form3.Show
End Sub
