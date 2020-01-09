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
      Caption         =   "Frame3"
      Height          =   1095
      Left            =   8040
      TabIndex        =   19
      Top             =   7920
      Width           =   6495
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "EXIT"
         Height          =   495
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "DATA MENU"
         Height          =   495
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "HOME MENU"
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
      Caption         =   "Frame2"
      Height          =   3015
      Left            =   8040
      TabIndex        =   18
      Top             =   4680
      Width           =   6495
      Begin VB.CommandButton Command8 
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
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   240
      TabIndex        =   4
      Top             =   4680
      Width           =   7575
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NEW"
         Height          =   495
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "BUY"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3720
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
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
      Begin MSAdodcLib.Adodc Adodc2 
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
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "DELETE"
         Height          =   495
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ADD"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text2 
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
   Begin MSDataGridLib.DataGrid CheckoutDataGrid 
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
   Begin MSAdodcLib.Adodc CheckoutAdodc 
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
   Begin VB.TextBox SearchText 
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

Private Sub CheckoutDataGrid_Click()
Text2.Text = CheckoutDataGrid.Columns(3)
Text3.Text = CheckoutDataGrid.Columns(4)
Text4.Text = CheckoutDataGrid.Columns(5)
Text5.Text = CheckoutDataGrid.Columns(6)
End Sub

Private Sub Command1_Click()
With Adodc2.Recordset
.AddNew
Adodc2.Recordset.Fields("MENU_ITEM") = Text2.Text
Adodc2.Recordset.Fields("QUANTITY") = Text3.Text
Adodc2.Recordset.Fields("PRICE") = Text4.Text
Adodc2.Recordset.Fields("TOTAL") = Text5.Text
Adodc2.Recordset.Update
Adodc2.RecordSource = "Select * FROM ORDERTABLE"
End With
End Sub

Private Sub Command2_Click()
If Adodc2.Recordset.RecordCount <> 0 Then Adodc2.Recordset.Delete
End Sub

Private Sub Command3_Click()
Adodc2.Recordset.MoveFirst
amount = 0
While Not Adodc2.Recordset.EOF
amount = amount + Adodc2.Recordset.Fields(3)
Adodc2.Recordset.MoveNext
Wend
TotalMoneyText.Text = amount
TotalMoneyText.Text = Format(amount, "###,##,0.00")
End Sub

Private Sub Command4_Click()
SearchText.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
TotalMoneyText.Text = ""
PaidMoneyText.Text = ""
ReturnMoneyText.Text = ""

Dim Mapus As Integer
For Mapus = 1 To Adodc2.Recordset.RecordCount
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Delete
Adodc2.Recordset.Update
Adodc2.Recordset.MoveNext
Next Mapus
End Sub

Private Sub Command6_Click()
FormDataMenu.Show
FormCashier.Hide
End Sub

Private Sub Command7_Click()
MsgBox "Are you sure you want to exit?", vbInformation + vbYesNo, "WARNING"
If vbYes Then End
End Sub

Private Sub Command8_Click()
Call PrintReceipt

End Sub

Private Sub Form_Load()
Call Connect_DB
CheckoutAdodc.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBCAFE.mdb"
CheckoutAdodc.RecordSource = "CHECKOUTTABLE"
CheckoutAdodc.Refresh
Set CheckoutDataGrid.DataSource = CheckoutAdodc

'Set AA = New ADODB.Connection
'AA.CursorLocation = adUseClient
'AA.Provider = "Microsoft.Jet.OLEDB.4.0"
'AA.Open App.Path & "\DBCAFE.mdb"
'Call AABB

Call Connect_DB
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBCAFE.mdb"
Adodc2.RecordSource = "ORDERTABLE"
Adodc2.Refresh
Set DataGrid2.DataSource = Adodc2

CheckoutDataGrid.Columns(0).Width = 700
CheckoutDataGrid.Columns(1).Width = 1500
CheckoutDataGrid.Columns(2).Width = 1900
CheckoutDataGrid.Columns(3).Width = 3500
CheckoutDataGrid.Columns(4).Width = 1300
CheckoutDataGrid.Columns(5).Width = 2000
CheckoutDataGrid.Columns(6).Width = 2000

DataGrid2.Columns(0).Width = 2300
DataGrid2.Columns(1).Width = 1300
DataGrid2.Columns(2).Width = 1300
DataGrid2.Columns(3).Width = 1300
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

Private Sub SearchText_Change()
Call Open_DB
RSSEARCH.Open "SELECT * FROM CHECKOUTTABLE WHERE TABLE_NUMBER like '%" & SearchText & "%'", Connection
If Not RSSEARCH.EOF Then
CheckoutAdodc.RecordSource = "SELECT * FROM CHECKOUTTABLE WHERE TABLE_NUMBER like '%" & SearchText & "%'"
CheckoutAdodc.Refresh
Set CheckoutDataGrid.DataSource = CheckoutAdodc

CheckoutDataGrid.Columns(0).Width = 700
CheckoutDataGrid.Columns(1).Width = 1500
CheckoutDataGrid.Columns(2).Width = 1900
CheckoutDataGrid.Columns(3).Width = 3500
CheckoutDataGrid.Columns(4).Width = 1300
CheckoutDataGrid.Columns(5).Width = 2000
CheckoutDataGrid.Columns(6).Width = 2000
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
CheckoutAdodc.RecordSource = "SELECT * from CHECKOUTTABLE WHERE TABLE_NUMBER = '" & CheckoutAdodc.Recordset!TABLE_NUMBER & "', Connection"
ReceiptScreen.Print Tab(5); "Table No.     :     "; CheckoutAdodc.Recordset!TABLE_NUMBER
ReceiptScreen.Print Tab(5); "Customer Name :     "; CheckoutAdodc.Recordset!CUSTOMER_NAME
MGrs = String$(33, "-")
ReceiptScreen.Print Tab(5); MGrs
Adodc2.RecordSource = "SELECT * FROM ORDERTABLE WHERE MENU_ITEM = '" & Adodc2.Recordset!MENU_ITEM & "', Connection"
Adodc2.Recordset.MoveFirst
No = 0
Do While Not Adodc2.Recordset.EOF
No = No + 1
Adodc2.RecordSource = "select * from ORDERTABLE WHERE MENU_ITEM = '" & Adodc2.Recordset!MENU_ITEM & "', Connection"
Quantity = Adodc2.Recordset!Quantity
Price = Adodc2.Recordset!Price
TOTAL = Quantity * Price
ReceiptScreen.Print Tab(5); No; Space(2); Adodc2.Recordset!MENU_ITEM
ReceiptScreen.Print Tab(10); DoRight(Quantity, "##"); Space(1); "X";
ReceiptScreen.Print Tab(15); Format(Price, "###,##,0.00");
ReceiptScreen.Print Tab(25); DoRight(TOTAL, "###,##,0.00");
Adodc2.Recordset.MoveNext
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
