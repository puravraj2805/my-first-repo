VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3030
   ClientLeft      =   525
   ClientTop       =   2220
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   Picture         =   "home.frx":0000
   Begin VB.Menu home 
      Caption         =   "HOME"
      WindowList      =   -1  'True
   End
   Begin VB.Menu tours 
      Caption         =   "TOURS"
   End
   Begin VB.Menu blog 
      Caption         =   "BLOG"
   End
   Begin VB.Menu aboutus 
      Caption         =   "ABOUT US"
   End
   Begin VB.Menu contactus 
      Caption         =   "CONTACT US"
   End
   Begin VB.Menu login 
      Caption         =   "LOGIN/SIGN UP"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub login_Click()
Unload Form1
Form1.Show
End Sub
