VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin MSAdodcLib.Adodc userado 
      Height          =   375
      Left            =   3000
      Top             =   9480
      Visible         =   0   'False
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\HP\Desktop\tour\user login.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\HP\Desktop\tour\user login.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from user"
      Caption         =   "user"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "SIGN IN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   13
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Frame user1 
      Caption         =   "USER LOGIN"
      Height          =   3975
      Left            =   5040
      TabIndex        =   6
      Top             =   3480
      Width           =   5415
      Begin VB.TextBox txtuser 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   600
         Width           =   3000
      End
      Begin VB.TextBox txtpassu 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   1560
         Width           =   3000
      End
      Begin VB.CommandButton cmd_loginusr 
         Caption         =   "LOGIN"
         Height          =   735
         Left            =   1680
         TabIndex        =   7
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label userid 
         Caption         =   "USER ID"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Width           =   1000
      End
      Begin VB.Label userpass 
         Caption         =   "PASSWORD"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1680
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
      Left            =   0
      TabIndex        =   5
      Top             =   2520
      Width           =   2000
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
      Left            =   2001
      TabIndex        =   4
      Top             =   2520
      Width           =   2000
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
      Left            =   4001
      TabIndex        =   3
      Top             =   2520
      Width           =   2000
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
      Left            =   6001
      TabIndex        =   2
      Top             =   2520
      Width           =   2000
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
      Left            =   8001
      TabIndex        =   1
      Top             =   2520
      Width           =   2000
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   0
      Picture         =   "userlogin.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   16155
      TabIndex        =   0
      Top             =   0
      Width           =   16215
   End
   Begin VB.Label Label1 
      Caption         =   "NOT A MEMBER?? "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   12
      Top             =   7920
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_loginusr_Click()
userado.RecordSource = "select * from user where userid='" + txtuser.Text + "' and userpass='" + txtpassu.Text + "'"
userado.Refresh
If userado.Recordset.EOF Then
    MsgBox "login failed,try again..!!!", vbCritical, "please enter correct user"
Else
    MsgBox "login successful.", vbInformation, "successful attempt"
    Form5.Show
End Sub

Private Sub Cmd1_Click()
Unload Form1
Form1.Show
End Sub

Private Sub Command1_Click()
Unload Form5
Form5.Show
End Sub
