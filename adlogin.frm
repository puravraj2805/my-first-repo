VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   9075
   ClientLeft      =   315
   ClientTop       =   -135
   ClientWidth     =   10245
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleMode       =   0  'User
   ScaleWidth      =   10245
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
      Left            =   16320
      TabIndex        =   11
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
      Left            =   12240
      TabIndex        =   10
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
      Left            =   8160
      TabIndex        =   9
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
      Left            =   4080
      TabIndex        =   8
      Top             =   3120
      Width           =   4059
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
      TabIndex        =   7
      Top             =   3120
      Width           =   4059
   End
   Begin MSAdodcLib.Adodc loginado 
      Height          =   495
      Left            =   6120
      Top             =   8400
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   873
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\HP\Desktop\tour\admin login.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\HP\Desktop\tour\admin login.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from admin"
      Caption         =   "login"
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
   Begin VB.Frame Frame1 
      Caption         =   "ADMIN LOGIN"
      Height          =   3975
      Left            =   7080
      TabIndex        =   1
      Top             =   3960
      Width           =   5415
      Begin VB.CommandButton cmd_login 
         Caption         =   "LOGIN"
         Height          =   735
         Left            =   1680
         TabIndex        =   6
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtpass 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1560
         Width           =   3000
      End
      Begin VB.TextBox txtadmin 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   600
         Width           =   3000
      End
      Begin VB.Label password 
         Caption         =   "PASSWORD"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1680
         Width           =   1000
      End
      Begin VB.Label adminid 
         Caption         =   "ADMIN ID"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   1000
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3135
      Left            =   0
      Picture         =   "adlogin.frx":0000
      ScaleHeight     =   3075
      ScaleWidth      =   20235
      TabIndex        =   0
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_login_Click()
loginado.RecordSource = "select * from admin where admin_id='" + txtadmin.Text + "' and pass='" + txtpass.Text + "'"
loginado.Refresh
If loginado.Recordset.EOF Then
    MsgBox "login failed,try again..!!!", vbCritical, "please enter correct user"
Else
    MsgBox "login successful.", vbInformation, "successful attempt"
    Form4.Show
End If
End Sub

Private Sub Cmd1_Click()
Unload Form1
Form1.Show
End Sub
