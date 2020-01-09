VERSION 5.00
Begin VB.Form FormLogin 
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      Height          =   615
      Left            =   3600
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Login to Gautam Cafe"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   1
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Label Label3 
      Caption         =   "PASSWORD"
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "NAME"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Conn As New ADODB.Connection
Public RS As New ADODB.Recordset

Sub Connection()
Set Conn = New ADODB.Connection
Set RS = New ADODB.Recordset
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBCAFE.mdb"
End Sub

Sub clear()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Command1_Click()
Call Connection
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "INCOMPLETE LOGIN DATA"
Exit Sub
Else
RS.Open "SELECT * FROM LOGINTABLE WHERE USER_NAME = '" & Text1 & "' AND PASSWORD = '" & Text2 & "'", Conn
If RS.EOF Then
MsgBox "INVALID CREDENTIALS!"
Call clear
Text1.SetFocus
Else
MsgBox "LOGIN SUCCEED!"
FormCashier.Show
FormLogin.Hide
End If
End If

End Sub

