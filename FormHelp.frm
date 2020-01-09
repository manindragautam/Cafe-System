VERSION 5.00
Begin VB.Form FormHelp 
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Back To Cafe Menu"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Back To Home"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   9255
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"FormHelp.frx":0000
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "HELP"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10995
   End
End
Attribute VB_Name = "FormHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FormHome.Show
FormHelp.Hide
End Sub

Private Sub Command2_Click()
FormMenu.Show
FormHelp.Hide
End Sub
