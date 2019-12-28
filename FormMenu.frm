VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormMenu 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   9510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   15345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameDessert 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Dessert Menu"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   360
      TabIndex        =   38
      Top             =   3840
      Width           =   9615
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   240
         TabIndex        =   43
         Text            =   "Select Dessert"
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1920
         TabIndex        =   42
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   1920
         TabIndex        =   41
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ADD"
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   6360
         TabIndex        =   39
         Top             =   4080
         Width           =   2895
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Lots"
         Height          =   495
         Left            =   240
         TabIndex        =   45
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Image Image3 
         Height          =   3135
         Left            =   4680
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PRICE (Rs.)"
         Height          =   375
         Left            =   4680
         TabIndex        =   44
         Top             =   4200
         Width           =   1935
      End
   End
   Begin VB.Frame FrameBeverage 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Beverage Menu"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   360
      TabIndex        =   29
      Top             =   3840
      Width           =   9615
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   240
         TabIndex        =   34
         Text            =   "Select Beverage"
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   1920
         TabIndex        =   33
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   1920
         TabIndex        =   32
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ADD"
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   6360
         TabIndex        =   30
         Top             =   4080
         Width           =   2895
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Lots"
         Height          =   495
         Left            =   240
         TabIndex        =   36
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Image Image2 
         Height          =   3135
         Left            =   4680
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PRICE (Rs.)"
         Height          =   375
         Left            =   4680
         TabIndex        =   35
         Top             =   4200
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Height          =   8295
      Left            =   10440
      TabIndex        =   22
      Top             =   360
      Width           =   4095
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "HELP"
         Height          =   615
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "HOME"
         Height          =   615
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   6960
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "RESERVE"
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   6960
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   1215
         Left            =   360
         TabIndex        =   25
         Top             =   5040
         Width           =   3375
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FormMenu.frx":0000
         Height          =   3495
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6165
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   360
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\shayank\Documents\GitHub\Cafe-System\DBCAFE.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\shayank\Documents\GitHub\Cafe-System\DBCAFE.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "CHECKOUTTABLE"
         Caption         =   "Adodc1"
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
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Order Total (Rs.)"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   24
         Top             =   4440
         Width           =   3375
      End
   End
   Begin VB.Frame FrameFood 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Food Menu"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   360
      TabIndex        =   13
      Top             =   3840
      Width           =   9615
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   6360
         TabIndex        =   21
         Top             =   4080
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ADD"
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Text            =   "Select Food"
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PRICE (Rs.)"
         Height          =   375
         Left            =   4680
         TabIndex        =   20
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   3135
         Left            =   4680
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Lots"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   2055
      Left            =   3960
      TabIndex        =   8
      Top             =   1440
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2280
         TabIndex        =   12
         Text            =   "Table Number"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   2280
         TabIndex        =   11
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Table Number         :"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name      :"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Menu Selection"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   3375
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   375
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dessert Menu"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Beverage Menu"
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Food Menu"
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":: CAFE MENU ::"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1920
      TabIndex        =   7
      Top             =   360
      Width           =   7215
   End
End
Attribute VB_Name = "FormMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()
FrameFood.Visible = True
FrameBeverage.Visible = False
FrameDessert.Visible = False
End Sub

Private Sub Option2_Click()
FrameFood.Visible = False
FrameBeverage.Visible = True
FrameDessert.Visible = False
End Sub

Private Sub Option3_Click()
FrameFood.Visible = False
FrameBeverage.Visible = False
FrameDessert.Visible = True
End Sub
