VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormDataMenu 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton UploadPictureButton 
      Caption         =   "UPLOAD PICTURE"
      Height          =   495
      Left            =   7560
      TabIndex        =   27
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Frame FrameFood 
      BackColor       =   &H00C0FFC0&
      Caption         =   "FOOD"
      Height          =   5175
      Left            =   360
      TabIndex        =   8
      Top             =   4200
      Width           =   9975
      Begin VB.CommandButton FoodDeleteButton 
         Caption         =   "DELETE"
         Height          =   615
         Left            =   5400
         TabIndex        =   13
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton FoodEditButton 
         Caption         =   "EDIT"
         Height          =   615
         Left            =   3960
         TabIndex        =   12
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton FoodSaveButton 
         Caption         =   "SAVE"
         Height          =   615
         Left            =   2520
         TabIndex        =   11
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "NEXT"
         Height          =   615
         Left            =   8400
         TabIndex        =   10
         Top             =   4080
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid FoodDataGrid 
         Height          =   3135
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5530
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   2415
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   6495
      Begin VB.CommandButton Command3 
         Caption         =   "CHECKOUT MENU"
         Height          =   495
         Left            =   1920
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton FormExitButton 
         Caption         =   "EXIT"
         Height          =   495
         Left            =   3480
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton DataNewButton 
         Caption         =   "NEW"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox TextPrice 
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox TextName 
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Price       :"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name      :"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame FrameBeverage 
      BackColor       =   &H00FFFFC0&
      Caption         =   "BEVERAGE"
      Height          =   5175
      Left            =   360
      TabIndex        =   14
      Top             =   4200
      Width           =   9975
      Begin VB.CommandButton Command12 
         Caption         =   "NEXT"
         Height          =   615
         Left            =   8400
         TabIndex        =   20
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton BeverageDeleteButton 
         Caption         =   "DELETE"
         Height          =   615
         Left            =   5400
         TabIndex        =   18
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton BeverageEditButton 
         Caption         =   "EDIT"
         Height          =   615
         Left            =   3960
         TabIndex        =   17
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton BeverageSaveButton 
         Caption         =   "SAVE"
         Height          =   615
         Left            =   2520
         TabIndex        =   16
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton Command11 
         Caption         =   "BACK"
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   4080
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid BeverageDataGrid 
         Height          =   3135
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5530
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
   End
   Begin VB.Frame FrameDessert 
      BackColor       =   &H00C0C0FF&
      Caption         =   "DESSERT"
      Height          =   5175
      Left            =   360
      TabIndex        =   21
      Top             =   4200
      Width           =   9975
      Begin VB.CommandButton DessertSaveButton 
         Caption         =   "SAVE"
         Height          =   615
         Left            =   2520
         TabIndex        =   25
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton DessertEditButton 
         Caption         =   "EDIT"
         Height          =   615
         Left            =   3960
         TabIndex        =   24
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton DessertDeleteButton 
         Caption         =   "DELETE"
         Height          =   615
         Left            =   5400
         TabIndex        =   23
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton Command13 
         Caption         =   "BACK"
         Height          =   615
         Left            =   240
         TabIndex        =   22
         Top             =   4080
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DessertDataGrid 
         Height          =   3495
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
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
   End
   Begin MSAdodcLib.Adodc DessertAdodc 
      Height          =   495
      Left            =   840
      Top             =   6120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
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
      Caption         =   "Adodc3"
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
      Height          =   495
      Left            =   840
      Top             =   5400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
      Height          =   495
      Left            =   840
      Top             =   4440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
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
   Begin MSComDlg.CommonDialog Upload 
      Left            =   12000
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "DATA MENU"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   28
      Top             =   240
      Width           =   6495
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3000
      Left            =   7320
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3000
   End
End
Attribute VB_Name = "FormDataMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim str As String
Dim stb As String
Dim Vimg As Boolean

Private Sub BeverageDataGrid_Click()
BeverageSaveButton.Enabled = False
TextName.Text = BeverageDataGrid.Columns(0)
TextPrice.Text = BeverageDataGrid.Columns(1)
stb = BeverageDataGrid.Columns(2)
Image1.Picture = LoadPicture(stb)
End Sub

Private Sub BeverageDeleteButton_Click()
MsgBox "Are you sure to delete this menu?", vbInformation + vbYesNo, "WARNING"
If vbYes Then
If BeverageAdodc.Recordset.RecordCount <> 0 Then BeverageAdodc.Recordset.Delete
Call clean
End If
End Sub

Private Sub BeverageEditButton_Click()
MsgBox "Are you sure to edit the data?", vbInformation + vbYesNo, "Menu Data"
If vbYes Then
BeverageAdodc.Recordset.Fields("BEVERAGE") = TextName.Text
BeverageAdodc.Recordset.Fields("PRICE") = TextPrice.Text
BeverageAdodc.Recordset.Fields("PICTURE") = str
BeverageAdodc.Recordset.Update
MsgBox "Data Edited Successfully", vbInformation, "Edit Menu"
End If
Call clean
End Sub

Private Sub BeverageSaveButton_Click()
With BeverageAdodc.Recordset
.AddNew
BeverageAdodc.Recordset.Fields("BEVERAGE") = TextName
BeverageAdodc.Recordset.Fields("PRICE") = TextPrice
BeverageAdodc.Recordset.Fields("PICTURE") = str
BeverageAdodc.Recordset.Update
MsgBox "New Menu Saved Successfully", vbInformation, "New Input Menu"
End With
End Sub

Private Sub Command3_Click()
FormCashier.Show
FormDataMenu.Hide
End Sub

Private Sub DataNewButton_Click()
Call clean
End Sub

Private Sub DessertDataGrid_Click()
DessertSaveButton.Enabled = False
TextName.Text = DessertDataGrid.Columns(0)
TextPrice.Text = DessertDataGrid.Columns(1)
stb = DessertDataGrid.Columns(2)
Image1.Picture = LoadPicture(stb)
End Sub

Private Sub DessertDeleteButton_Click()
MsgBox "Are you sure to delete this menu?", vbInformation + vbYesNo, "WARNING"
If vbYes Then
If DessertAdodc.Recordset.RecordCount <> 0 Then DessertAdodc.Recordset.Delete
Call clean
End If
End Sub

Private Sub DessertEditButton_Click()
MsgBox "Are you sure to edit the data?", vbInformation + vbYesNo, "Menu Data"
If vbYes Then
DessertAdodc.Recordset.Fields("DESSERT") = TextName.Text
DessertAdodc.Recordset.Fields("PRICE") = TextPrice.Text
DessertAdodc.Recordset.Fields("PICTURE") = str
DessertAdodc.Recordset.Update
MsgBox "Data Edited Successfully", vbInformation, "Edit Menu"
End If
Call clean
End Sub

Private Sub DessertSaveButton_Click()
With DessertAdodc.Recordset
.AddNew
DessertAdodc.Recordset.Fields("DESSERT") = TextName
DessertAdodc.Recordset.Fields("PRICE") = TextPrice
DessertAdodc.Recordset.Fields("PICTURE") = str
DessertAdodc.Recordset.Update
MsgBox "New Menu Saved Successfully", vbInformation, "New Input Menu"
End With
End Sub

Private Sub FoodDeleteButton_Click()
MsgBox "Are you sure to delete this menu?", vbInformation + vbYesNo, "WARNING"
If vbYes Then
If FoodAdodc.Recordset.RecordCount <> 0 Then FoodAdodc.Recordset.Delete
Call clean
End If
End Sub

Private Sub FoodEditButton_Click()
MsgBox "Are you sure to edit the data?", vbInformation + vbYesNo, "Menu Data"
If vbYes Then
FoodAdodc.Recordset.Fields("FOOD") = TextName.Text
FoodAdodc.Recordset.Fields("PRICE") = TextPrice.Text
FoodAdodc.Recordset.Fields("PICTURE") = str
FoodAdodc.Recordset.Update
MsgBox "Data Edited Successfully", vbInformation, "Edit Menu"
End If
Call clean
End Sub

Private Sub FoodSaveButton_Click()
With FoodAdodc.Recordset
.AddNew
FoodAdodc.Recordset.Fields("FOOD") = TextName
FoodAdodc.Recordset.Fields("PRICE") = TextPrice
FoodAdodc.Recordset.Fields("PICTURE") = str
FoodAdodc.Recordset.Update
MsgBox "New Menu Saved Successfully", vbInformation, "New Input Menu"
End With
End Sub

Private Sub Command11_Click()
FrameBeverage.Visible = False
FrameFood.Visible = True
End Sub

Private Sub Command12_Click()
FrameDessert.Visible = True
FrameBeverage.Visible = False
End Sub

Private Sub Command13_Click()
FrameBeverage.Visible = True
FrameDessert.Visible = False
End Sub

Private Sub Command5_Click()
FrameBeverage.Visible = True
FrameFood.Visible = False
End Sub

Private Sub FoodDataGrid_Click()
FoodSaveButton.Enabled = False
TextName.Text = FoodDataGrid.Columns(0)
TextPrice.Text = FoodDataGrid.Columns(1)
stb = FoodDataGrid.Columns(2)
Image1.Picture = LoadPicture(stb)
End Sub

Private Sub Form_Load()
Call Connect_DB
FoodAdodc.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBCAFE.mdb"
FoodAdodc.RecordSource = "FOODTABLE"
FoodAdodc.Refresh
Set FoodDataGrid.DataSource = FoodAdodc

Call Connect_DB
BeverageAdodc.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBCAFE.mdb"
BeverageAdodc.RecordSource = "BEVERAGETABLE"
BeverageAdodc.Refresh
Set BeverageDataGrid.DataSource = BeverageAdodc

Call Connect_DB
DessertAdodc.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBCAFE.mdb"
DessertAdodc.RecordSource = "DESSERTTABLE"
DessertAdodc.Refresh
Set DessertDataGrid.DataSource = DessertAdodc

End Sub

Sub clean()
TextName.Text = ""
TextPrice.Text = ""
Image1.Picture = LoadPicture("")
FoodSaveButton.Enabled = True
BeverageSaveButton.Enabled = True
DessertSaveButton.Enabled = True
End Sub

Private Sub FormExitButton_Click()
MsgBox "Are you sure you want to exit?", vbInformation + vbYesNo, "WARNING"
If vbYes Then End
End Sub

Private Sub UploadPictureButton_Click()
Upload.ShowOpen
Upload.Filter = "Jpeg|*.jpg"
str = Upload.FileName
Image1.Picture = LoadPicture(str)
Vimg = True
End Sub
