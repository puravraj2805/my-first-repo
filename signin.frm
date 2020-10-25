VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   9120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame user1 
      Caption         =   "USER REGISTERATION"
      Height          =   3975
      Left            =   8160
      TabIndex        =   6
      Top             =   4560
      Width           =   5415
      Begin VB.CommandButton cmd_loginusr 
         Caption         =   "SIGN IN"
         Height          =   735
         Left            =   1800
         TabIndex        =   9
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtpassu 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   1560
         Width           =   3000
      End
      Begin VB.TextBox txtuser 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   600
         Width           =   3000
      End
      Begin VB.Label userpass 
         Caption         =   "PASSWORD"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1680
         Width           =   1000
      End
      Begin VB.Label userid 
         Caption         =   "USER ID"
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   720
         Width           =   1000
      End
   End
   Begin VB.CommandButton Cmd1 
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
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   4059
   End
   Begin VB.CommandButton Cmd2 
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
      Left            =   4200
      TabIndex        =   4
      Top             =   3120
      Width           =   4059
   End
   Begin VB.CommandButton Cmd3 
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
      Left            =   8280
      TabIndex        =   3
      Top             =   3120
      Width           =   4059
   End
   Begin VB.CommandButton Cmd4 
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
      Left            =   12360
      TabIndex        =   2
      Top             =   3120
      Width           =   4059
   End
   Begin VB.CommandButton Cmd5 
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
      Left            =   16440
      TabIndex        =   1
      Top             =   3120
      Width           =   4059
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3135
      Left            =   0
      Picture         =   "signin.frx":0000
      ScaleHeight     =   3075
      ScaleWidth      =   20235
      TabIndex        =   0
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d As ADODB.Connection
Dim r As ADODB.Recordset
Dim s As String

Private Sub cmd_loginusr_Click()
r.AddNew
End Sub

Private Sub Form_Load()
Set d = New ADODB.Connection
d.ConnectionString = "provider=MICROSOFT.JET.OLEDB.4.0;DataSource = C:\Users\HP\Desktop\tour\user login.mdb;"
d.Open
MsgBox "database open"
Set r = New ADODB.Recordset
s = "select * from user"
r.Open s, d, adOpenDynamic, adLockOptimistic
End Sub
