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
      Begin VB.CommandButton ReserveButton 
         BackColor       =   &H00C0FFFF&
         Caption         =   "RESERVE"
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   6960
         Width           =   975
      End
      Begin VB.TextBox OrderTotalText 
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   360
         TabIndex        =   25
         Top             =   5040
         Width           =   3375
      End
      Begin MSDataGridLib.DataGrid CheckoutDataGrid 
         Bindings        =   "FormMenu.frx":0000
         Height          =   3495
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
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
      Begin MSAdodcLib.Adodc CheckoutAdodc 
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
      Begin VB.TextBox FoodTotalPriceText 
         Height          =   495
         Left            =   6360
         TabIndex        =   21
         Top             =   4080
         Width           =   2895
      End
      Begin VB.CommandButton FoodAddButton 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ADD"
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox FoodQuantityText 
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox FoodPriceText 
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox FoodSelector 
         Height          =   315
         ItemData        =   "FormMenu.frx":001C
         Left            =   240
         List            =   "FormMenu.frx":001E
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
      Begin VB.Image FoodImage 
         Height          =   3600
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   600
         Width           =   3600
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
      Begin VB.ComboBox TableNumberSelector 
         Height          =   315
         ItemData        =   "FormMenu.frx":0020
         Left            =   2280
         List            =   "FormMenu.frx":0042
         TabIndex        =   12
         Text            =   "Table Number"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox CustomerNameText 
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
      Begin VB.OptionButton DessertRadioSwitch 
         BackColor       =   &H00C0E0FF&
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   375
      End
      Begin VB.OptionButton BeverageRadioSwitch 
         BackColor       =   &H00C0E0FF&
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   375
      End
      Begin VB.OptionButton FoodRadioSwitch 
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
      Begin VB.ComboBox BeverageSelector 
         Height          =   315
         Left            =   240
         TabIndex        =   34
         Text            =   "Select Beverage"
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox BeveragePriceText 
         Height          =   375
         Left            =   1920
         TabIndex        =   33
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox BeverageQuantityText 
         Height          =   375
         Left            =   1920
         TabIndex        =   32
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CommandButton BeverageAddButton 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ADD"
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox BeverageTotalPriceText 
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
         Caption         =   "Quantity"
         Height          =   495
         Left            =   240
         TabIndex        =   36
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Image BeverageImage 
         Height          =   3600
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   600
         Width           =   3600
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
      Begin VB.TextBox DessertTotalPriceText 
         Height          =   495
         Left            =   6360
         TabIndex        =   43
         Top             =   4080
         Width           =   2895
      End
      Begin VB.CommandButton DessertAddButton 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ADD"
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox DessertQuantityText 
         Height          =   375
         Left            =   1920
         TabIndex        =   41
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox DessertPriceText 
         Height          =   375
         Left            =   1920
         TabIndex        =   40
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox DessertSelector 
         Height          =   315
         Left            =   240
         TabIndex        =   39
         Text            =   "Select Dessert"
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PRICE (Rs.)"
         Height          =   375
         Left            =   4680
         TabIndex        =   46
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Image DessertImage 
         Height          =   3600
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   600
         Width           =   3600
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   495
         Left            =   240
         TabIndex        =   45
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   1200
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc DessertAdodc 
      Height          =   375
      Left            =   720
      Top             =   5400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc BeverageAdodc 
      Height          =   330
      Left            =   720
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc FoodAdodc 
      Height          =   375
      Left            =   720
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
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
Private Sub BeverageAddButton_Click()
If CustomerNameText.Text = "" Or TableNumberSelector.Text = "Table Number" Or BeverageSelector.Text = "" Or BeveragePriceText.Text = "" Or BeverageQuantityText.Text = "" Then
MsgBox " PLEASE COMPLETE YOUR ORDER ", vbInformation, "WARNING"
Else

With CheckoutAdodc.Recordset
.AddNew
CheckoutAdodc.Recordset.Fields("TABLE_NUMBER") = TableNumberSelector.Text
CheckoutAdodc.Recordset.Fields("CUSTOMER_NAME") = CustomerNameText.Text
CheckoutAdodc.Recordset.Fields("ORDER") = BeverageSelector.Text
CheckoutAdodc.Recordset.Fields("PRICE") = BeveragePriceText.Text
CheckoutAdodc.Recordset.Fields("QUANTITY") = BeverageQuantityText.Text
CheckoutAdodc.Recordset.Fields("TOTAL") = BeverageTotalPriceText.Text
OrderTotalText.Text = Val(OrderTotalText.Text) + Val(BeverageTotalPriceText.Text)
CheckoutAdodc.Recordset.Update
CheckoutAdodc.RecordSource = "Select * From CHECKOUTTABLE"
MsgBox " Data Already Saved "

BeverageQuantityText.Text = ""
BeveragePriceText.Text = ""
BeverageTotalPriceText.Text = ""
BeverageSelector.Text = "Select Beverage"
BeverageImage.Picture = LoadPicture("")
End With
End If
End Sub

Private Sub BeverageQuantityText_Change()
BeverageTotalPriceText.Text = Val(BeverageQuantityText.Text) * Val(BeveragePriceText.Text)
End Sub

Private Sub BeverageRadioSwitch_Click()
FrameFood.Visible = False
FrameBeverage.Visible = True
FrameDessert.Visible = False
End Sub

Private Sub BeverageSelector_Click()
BeverageAdodc.Recordset.Filter = " BEVERAGE = '" & BeverageSelector & "'"
BeveragePriceText.Text = BeverageAdodc.Recordset!PRICE
stb = BeverageAdodc.Recordset!Picture
BeverageImage.Picture = LoadPicture(stb)
End Sub


Private Sub DessertAddButton_Click()
If CustomerNameText.Text = "" Or TableNumberSelector.Text = "Table Number" Or DessertSelector.Text = "" Or DessertPriceText.Text = "" Or DessertQuantityText.Text = "" Then
MsgBox " PLEASE COMPLETE YOUR ORDER ", vbInformation, "WARNING"
Else

With CheckoutAdodc.Recordset
.AddNew
CheckoutAdodc.Recordset.Fields("TABLE_NUMBER") = TableNumberSelector.Text
CheckoutAdodc.Recordset.Fields("CUSTOMER_NAME") = CustomerNameText.Text
CheckoutAdodc.Recordset.Fields("ORDER") = DessertSelector.Text
CheckoutAdodc.Recordset.Fields("PRICE") = DessertPriceText.Text
CheckoutAdodc.Recordset.Fields("QUANTITY") = DessertQuantityText.Text
CheckoutAdodc.Recordset.Fields("TOTAL") = DessertTotalPriceText.Text
OrderTotalText.Text = Val(OrderTotalText.Text) + Val(DessertTotalPriceText.Text)
CheckoutAdodc.Recordset.Update
CheckoutAdodc.RecordSource = "Select * From CHECKOUTTABLE"
MsgBox " Data Already Saved "

DessertQuantityText.Text = ""
DessertPriceText.Text = ""
DessertTotalPriceText.Text = ""
DessertSelector.Text = "Select Dessert"
DessertImage.Picture = LoadPicture("")
End With
End If
End Sub

Private Sub DessertQuantityText_Change()
DessertTotalPriceText.Text = Val(DessertQuantityText.Text) * Val(DessertPriceText.Text)
End Sub

Private Sub DessertRadioSwitch_Click()
FrameFood.Visible = False
FrameBeverage.Visible = False
FrameDessert.Visible = True
End Sub

Private Sub DessertSelector_Click()
DessertAdodc.Recordset.Filter = " DESSERT = '" & DessertSelector & "'"
DessertPriceText.Text = DessertAdodc.Recordset!PRICE
stb = DessertAdodc.Recordset!Picture
DessertImage.Picture = LoadPicture(stb)
End Sub

Private Sub FoodAddButton_Click()
If CustomerNameText.Text = "" Or TableNumberSelector.Text = "Table Number" Or FoodSelector.Text = "" Or FoodPriceText.Text = "" Or FoodQuantityText.Text = "" Then
MsgBox " PLEASE COMPLETE YOUR ORDER ", vbInformation, "WARNING"
Else

With CheckoutAdodc.Recordset
.AddNew
CheckoutAdodc.Recordset.Fields("TABLE_NUMBER") = TableNumberSelector.Text
CheckoutAdodc.Recordset.Fields("CUSTOMER_NAME") = CustomerNameText.Text
CheckoutAdodc.Recordset.Fields("ORDER") = FoodSelector.Text
CheckoutAdodc.Recordset.Fields("PRICE") = FoodPriceText.Text
CheckoutAdodc.Recordset.Fields("QUANTITY") = FoodQuantityText.Text
CheckoutAdodc.Recordset.Fields("TOTAL") = FoodTotalPriceText.Text
OrderTotalText.Text = Val(OrderTotalText.Text) + Val(FoodTotalPriceText.Text)
CheckoutAdodc.Recordset.Update
CheckoutAdodc.RecordSource = "Select * From CHECKOUTTABLE"
MsgBox " Data Already Saved "

FoodQuantityText.Text = ""
FoodPriceText.Text = ""
FoodTotalPriceText.Text = ""
FoodSelector.Text = "Select Food"
FoodImage.Picture = LoadPicture("")
End With
End If
End Sub

Private Sub FoodQuantityText_Change()
FoodTotalPriceText.Text = Val(FoodQuantityText.Text) * Val(FoodPriceText.Text)
End Sub

Private Sub FoodRadioSwitch_Click()
FrameFood.Visible = True
FrameBeverage.Visible = False
FrameDessert.Visible = False
End Sub

Private Sub FoodSelector_Click()
FoodAdodc.Recordset.Filter = " FOOD = '" & FoodSelector & "'"
FoodPriceText.Text = FoodAdodc.Recordset!PRICE
stb = FoodAdodc.Recordset!Picture
FoodImage.Picture = LoadPicture(stb)
End Sub

Private Sub Form_Load()
CheckoutDataGrid.Columns(0).Width = 500
CheckoutDataGrid.Columns(1).Width = 1400
CheckoutDataGrid.Columns(2).Width = 1500
CheckoutDataGrid.Columns(3).Width = 2000
CheckoutDataGrid.Columns(4).Width = 1000
CheckoutDataGrid.Columns(5).Width = 1000
CheckoutDataGrid.Columns(6).Width = 1000
End Sub

Private Sub Form_Activate()
Call Connect_DB
FoodAdodc.ConnectionString = Connection.ConnectionString
FoodAdodc.RecordSource = "Select * From FOODTABLE"
FoodAdodc.Refresh
FoodSelector.Clear
FoodSelector = "Select Food"
With FoodAdodc.Recordset
Do While Not .EOF
FoodSelector.AddItem !FOOD
FoodAdodc.Recordset.MoveNext
Loop
End With

Call Connect_DB
BeverageAdodc.ConnectionString = Connection.ConnectionString
BeverageAdodc.RecordSource = "Select * From BEVERAGETABLE"
BeverageAdodc.Refresh
BeverageSelector.Clear
BeverageSelector = "Select Beverage"
With BeverageAdodc.Recordset
Do While Not .EOF
BeverageSelector.AddItem !BEVERAGE
BeverageAdodc.Recordset.MoveNext
Loop
End With

Call Connect_DB
DessertAdodc.ConnectionString = Connection.ConnectionString
DessertAdodc.RecordSource = "Select * From DESSERTTABLE"
DessertAdodc.Refresh
DessertSelector.Clear
DessertSelector = "Select Dessert"
With DessertAdodc.Recordset
Do While Not .EOF
DessertSelector.AddItem !DESSERT
DessertAdodc.Recordset.MoveNext
Loop
End With
End Sub

Private Sub ReserveButton_Click()
If CustomerNameText.Text = "" Or TableNumberSelector.Text = "Table Number" Or OrderTotalText.Text = "" Then
MsgBox " PLEASE COMPLETE YOUR ORDER ", vbInformation, "WARNING"
Else
MsgBox " Thanks, Your order is in process, Please proceed the payment at checkout ^_^ "
CustomerNameText.Text = ""
TableNumberSelector.Text = "Table Number"
OrderTotalText.Text = ""

FoodImage.Picture = LoadPicture()
BeverageImage.Picture = LoadPicture()
DessertImage.Picture = LoadPicture()
End If
End Sub
