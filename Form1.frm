VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML // Data Box by: -Theory-"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "Horizontal Line"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Line Break"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Change Font"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Hyperlink Image"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Display Image"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Change Page Title"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hyperlink in a New Window"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Make a Hyperlink"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change Background Image"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Background Color"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "The Code"
      ForeColor       =   &H0000FF00&
      Height          =   4815
      Left            =   2760
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.TextBox Text1 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   4455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "Form1.frx":0000
         Top             =   240
         Width           =   5055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MADE BY: -Theory-

Private Sub Command1_Click()
Text1.Text = "<body bgcolor=""COLORHERE"">"
End Sub

Private Sub Command10_Click()
Text1.Text = "<hr>"
End Sub

Private Sub Command2_Click()
Text1.Text = "<body backcolor=""imageurl.gif"">"
End Sub

Private Sub Command3_Click()
Text1.Text = "<a href=""http://yoursite.com"">Page Name</a>"
End Sub

Private Sub Command4_Click()
Text1.Text = "<a href=""http://yoursite.com"" target=""new"">Page Name</a>"
End Sub

Private Sub Command5_Click()
Text1.Text = "<title>YOUR TITLE HERE</title>"
End Sub

Private Sub Command6_Click()
Text1.Text = "<img src=""imgname.gif"">"
End Sub

Private Sub Command7_Click()
Text1.Text = "<a href=""http://yoursite.com""><img src=""imagename.gif""></a>"
End Sub

Private Sub Command8_Click()
Text1.Text = "<font name=""verdana"" size=""1"" color=""COLORHERE"">"
End Sub

Private Sub Command9_Click()
Text1.Text = "<br>"
End Sub
