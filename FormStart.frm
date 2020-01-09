VERSION 5.00
Begin VB.Form FormStart 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CASHIER"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CUSTOMER"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   2160
      TabIndex        =   1
      Top             =   2280
      Width           =   5895
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "SELECT YOUR CHOICE"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "GAUTAM CAFE"
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
Attribute VB_Name = "FormStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
FormHome.Show
FormStart.Hide
End Sub
