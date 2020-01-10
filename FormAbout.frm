VERSION 5.00
Begin VB.Form FormAbout 
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BackButton 
      BackColor       =   &H008080FF&
      Caption         =   "BACK"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   9255
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"FormAbout.frx":0000
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
      Caption         =   "ABOUT"
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
      Top             =   240
      Width           =   10995
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BackButton_Click()
FormHome.Show
FormAbout.Hide
End Sub

