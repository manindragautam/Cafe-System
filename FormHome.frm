VERSION 5.00
Begin VB.Form FormHome 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   10575
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Eat Fresh, Delicious Food"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   735
         Left            =   1200
         TabIndex        =   6
         Top             =   5160
         Width           =   7815
      End
      Begin VB.Image Image1 
         Height          =   4920
         Left            =   600
         Picture         =   "FormHome.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   9240
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "HELP"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "CAFE MENU"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "ABOUT"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "WELCOME TO GAUTAM CAFE"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   -120
      TabIndex        =   0
      Top             =   240
      Width           =   10995
   End
End
Attribute VB_Name = "FormHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FormAbout.Show
FormHome.Hide
End Sub

Private Sub Command2_Click()
FormMenu.Show
FormHome.Hide
End Sub

Private Sub Command3_Click()
FormHelp.Show
FormHome.Hide
End Sub

Private Sub Command4_Click()
FormStart.Show
FormHome.Hide
End Sub

