VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormCashier 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Height          =   1095
      Left            =   8040
      TabIndex        =   19
      Top             =   7920
      Width           =   6495
      Begin VB.CommandButton ExitButton 
         BackColor       =   &H00C0E0FF&
         Caption         =   "EXIT"
         Height          =   495
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton CafeMenuButton 
         BackColor       =   &H00C0E0FF&
         Caption         =   "CAFE MENU"
         Height          =   495
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton HomeButton 
         BackColor       =   &H00C0E0FF&
         Caption         =   "HOME"
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   3015
      Left            =   8040
      TabIndex        =   18
      Top             =   4680
      Width           =   6495
      Begin VB.CommandButton PrintButton 
         BackColor       =   &H00C0FFFF&
         Caption         =   "PRINT RECEIPT"
         Height          =   495
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox ReturnMoneyText 
         Height          =   375
         Left            =   3240
         TabIndex        =   25
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox PaidMoneyText 
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox TotalMoneyText 
         Height          =   375
         Left            =   3240
         TabIndex        =   23
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "RETURN (Rs.)"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   375
         Left            =   600
         TabIndex        =   28
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "PAID (Rs.)"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   600
         TabIndex        =   27
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL (Rs.)"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   600
         TabIndex        =   26
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   4335
      Left            =   240
      TabIndex        =   4
      Top             =   4680
      Width           =   7575
      Begin VB.CommandButton RefreshButton 
         BackColor       =   &H00C0FFFF&
         Caption         =   "REFRESH"
         Height          =   495
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton CheckoutButton 
         BackColor       =   &H00C0FFFF&
         Caption         =   "CHECKOUT"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3720
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid CheckoutDataGrid 
         Bindings        =   "FormCashier.frx":0000
         Height          =   1575
         Left            =   360
         TabIndex        =   15
         Top             =   2040
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   2778
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
         Left            =   840
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.CommandButton DeleteButton 
         BackColor       =   &H00C0FFFF&
         Caption         =   "DELETE"
         Height          =   495
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton AddButton 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ADD"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox TotalText 
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox PriceText 
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox QuantityText 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox OrderText 
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
         Height          =   375
         Left            =   5040
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRICE"
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "QUANTITY"
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ORDER"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid TableDataGrid 
      Height          =   3135
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   14295
      _ExtentX        =   25215
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
   Begin MSAdodcLib.Adodc TableAdodc 
      Height          =   495
      Left            =   1200
      Top             =   2280
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.TextBox TableSearchText 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "CASHIER"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   5640
      TabIndex        =   2
      Top             =   240
      Width           =   8295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Table Number"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "FormCashier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RSCHECKOUT As New ADODB.Recordset
Public RSORDER As New ADODB.Recordset
Dim Connection As New ADODB.Connection
Dim RSSEARCH As New ADODB.Recordset

Private Sub HomeButton_Click()
FormHome.Show
FormCashier.Hide
End Sub

Private Sub TableDataGrid_Click()
OrderText.Text = TableDataGrid.Columns(3)
QuantityText.Text = TableDataGrid.Columns(4)
PriceText.Text = TableDataGrid.Columns(5)
TotalText.Text = TableDataGrid.Columns(6)
End Sub

Private Sub AddButton_Click()
With CheckoutAdodc.Recordset
.AddNew
CheckoutAdodc.Recordset.Fields("MENU_ITEM") = OrderText.Text
CheckoutAdodc.Recordset.Fields("QUANTITY") = QuantityText.Text
CheckoutAdodc.Recordset.Fields("PRICE") = PriceText.Text
CheckoutAdodc.Recordset.Fields("TOTAL") = TotalText.Text
CheckoutAdodc.Recordset.Update
CheckoutAdodc.RecordSource = "Select * FROM ORDERTABLE"
End With
End Sub

Private Sub DeleteButton_Click()
If CheckoutAdodc.Recordset.RecordCount <> 0 Then CheckoutAdodc.Recordset.Delete
End Sub

Private Sub CheckoutButton_Click()
CheckoutAdodc.Recordset.MoveFirst
amount = 0
While Not CheckoutAdodc.Recordset.EOF
amount = amount + CheckoutAdodc.Recordset.Fields(3)
CheckoutAdodc.Recordset.MoveNext
Wend
TotalMoneyText.Text = amount
TotalMoneyText.Text = Format(amount, "###,##,0.00")
End Sub

Private Sub RefreshButton_Click()
TableSearchText.Text = ""
OrderText.Text = ""
QuantityText.Text = ""
PriceText.Text = ""
TotalText.Text = ""
TotalMoneyText.Text = ""
PaidMoneyText.Text = ""
ReturnMoneyText.Text = ""

Dim InitNum As Integer
For InitNum = 1 To CheckoutAdodc.Recordset.RecordCount
CheckoutAdodc.Recordset.MoveFirst
CheckoutAdodc.Recordset.Delete
CheckoutAdodc.Recordset.Update
CheckoutAdodc.Recordset.MoveNext
Next InitNum
End Sub

Private Sub CafeMenuButton_Click()
FormCafeMenu.Show
FormCashier.Hide
End Sub

Private Sub ExitButton_Click()
MsgBox "Are you sure you want to exit?", vbInformation + vbYesNo, "WARNING"
If vbYes Then End
End Sub

Private Sub PrintButton_Click()
Call PrintReceipt
End Sub

Private Sub Form_Load()
Call Connect_DB
TableAdodc.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBCAFE.mdb"
TableAdodc.RecordSource = "CHECKOUTTABLE"
TableAdodc.Refresh
Set TableDataGrid.DataSource = TableAdodc

Call Connect_DB
CheckoutAdodc.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBCAFE.mdb"
CheckoutAdodc.RecordSource = "ORDERTABLE"
CheckoutAdodc.Refresh
Set CheckoutDataGrid.DataSource = CheckoutAdodc

TableDataGrid.Columns(0).Width = 700
TableDataGrid.Columns(1).Width = 1500
TableDataGrid.Columns(2).Width = 1900
TableDataGrid.Columns(3).Width = 3500
TableDataGrid.Columns(4).Width = 1300
TableDataGrid.Columns(5).Width = 2000
TableDataGrid.Columns(6).Width = 2000

CheckoutDataGrid.Columns(0).Width = 2300
CheckoutDataGrid.Columns(1).Width = 1300
CheckoutDataGrid.Columns(2).Width = 1300
CheckoutDataGrid.Columns(3).Width = 1300
End Sub

Sub Open_DB()
Set Connection = New ADODB.Connection
Set RSSEARCH = New ADODB.Recordset
Connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBCAFE.mdb"
End Sub

Private Sub PaidMoneyText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TOTAL = TotalMoneyText.Text
PAID = PaidMoneyText.Text
PaidMoneyText.Text = Format(PAID, "###,##,0.00")
RETURNCASH = PAID - TOTAL
ReturnMoneyText.Text = Format(RETURNCASH, "###,##,0.00")
End If
End Sub

Private Sub TableSearchText_Change()
Call Open_DB
RSSEARCH.Open "SELECT * FROM CHECKOUTTABLE WHERE TABLE_NUMBER like '%" & TableSearchText & "%'", Connection
If Not RSSEARCH.EOF Then
TableAdodc.RecordSource = "SELECT * FROM CHECKOUTTABLE WHERE TABLE_NUMBER like '%" & TableSearchText & "%'"
TableAdodc.Refresh
Set TableDataGrid.DataSource = TableAdodc

TableDataGrid.Columns(0).Width = 700
TableDataGrid.Columns(1).Width = 1500
TableDataGrid.Columns(2).Width = 1900
TableDataGrid.Columns(3).Width = 3500
TableDataGrid.Columns(4).Width = 1300
TableDataGrid.Columns(5).Width = 2000
TableDataGrid.Columns(6).Width = 2000
End If
End Sub

Sub dbconnection()
Set Connection = New ADODB.Connection
Set RSCHECKOUT = New ADODB.Recordset
Set RSORDER = New ADODB.Recordset
Set RSSEARCH = New ADODB.Recordset
Connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBCAFE.mdb"
End Sub

Function PrintReceipt()
Call dbconnection
ReceiptScreen.Show
Dim TotalMoney, PaidMoney, ReturnMoney As Double
Dim MGrs As String
ReceiptScreen.Font = "Courier New"
ReceiptScreen.Print
ReceiptScreen.Print
ReceiptScreen.Print Tab(15); "GAUTAM CAFE";
ReceiptScreen.Print Tab(20); "RECEIPT";
ReceiptScreen.Print
ReceiptScreen.Print
ReceiptScreen.Print
ReceiptScreen.Print
ReceiptScreen.Print
TableAdodc.RecordSource = "SELECT * from CHECKOUTTABLE WHERE TABLE_NUMBER = '" & TableAdodc.Recordset!TABLE_NUMBER & "', Connection"
ReceiptScreen.Print Tab(5); "Table No.     :     "; TableAdodc.Recordset!TABLE_NUMBER
ReceiptScreen.Print Tab(5); "Customer Name :     "; TableAdodc.Recordset!CUSTOMER_NAME
MGrs = String$(33, "-")
ReceiptScreen.Print Tab(5); MGrs
CheckoutAdodc.Recordset.MoveFirst
No = 0
Do While Not CheckoutAdodc.Recordset.EOF
No = No + 1
CheckoutAdodc.RecordSource = "select * from ORDERTABLE WHERE MENU_ITEM = '" & CheckoutAdodc.Recordset!MENU_ITEM & "', Connection"
Quantity = CheckoutAdodc.Recordset!Quantity
Price = CheckoutAdodc.Recordset!Price
TOTAL = Quantity * Price
ReceiptScreen.Print Tab(5); No; Space(2); CheckoutAdodc.Recordset!MENU_ITEM
ReceiptScreen.Print Tab(10); DoRight(Quantity, "##"); Space(1); "X";
ReceiptScreen.Print Tab(15); Format(Price, "###,##,0.00");
ReceiptScreen.Print Tab(25); DoRight(TOTAL, "###,##,0.00");
CheckoutAdodc.Recordset.MoveNext
Loop

ReceiptScreen.Print Tab(5); MGrs
ReceiptScreen.Print Tab(5); "Total         :";
ReceiptScreen.Print Tab(25); DoRight(TotalMoneyText.Text, "###,##,0.00");
ReceiptScreen.Print Tab(5); "Paid          :";
ReceiptScreen.Print Tab(25); DoRight(PaidMoneyText.Text, "###,##,0.00");
ReceiptScreen.Print Tab(5); MGrs
ReceiptScreen.Print Tab(5); "Return        :";
ReceiptScreen.Print Tab(25); DoRight(ReturnMoneyText.Text, "###,##,0.00");
ReceiptScreen.Print Tab(5); MGrs
ReceiptScreen.Print Tab(6); "Thank you for your visit"
ReceiptScreen.Print
ReceiptScreen.Print
ReceiptScreen.Print
ReceiptScreen.Print
End Function

Private Function DoRight(NData, CFormat) As String
DoRight = Format(NData, CFormat)
DoRight = Space(Len(CFormat) - Len(DoRight)) + DoRight
End Function
