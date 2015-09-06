VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form logging_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Add Record"
   ClientHeight    =   7275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   14370
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton delorder_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Delete Order"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton addorder_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Add Order"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton process_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cancel_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton increment_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton decrement_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox stockname_txtbox 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid orders_datgrid 
      Height          =   2985
      Left            =   0
      TabIndex        =   20
      Top             =   2160
      Width           =   14040
      _ExtentX        =   24765
      _ExtentY        =   5265
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777088
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   29
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15
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
   Begin VB.Label orderdate_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   13560
      TabIndex        =   24
      Top             =   600
      Width           =   45
   End
   Begin VB.Label orderdate_lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12360
      TabIndex        =   23
      Top             =   600
      Width           =   945
   End
   Begin VB.Label orderno_lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order No:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12360
      TabIndex        =   22
      Top             =   360
      Width           =   795
   End
   Begin VB.Label orderno_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   13560
      TabIndex        =   21
      Top             =   360
      Width           =   45
   End
   Begin VB.Label totalprice_lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Price:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3960
      TabIndex        =   19
      Top             =   5880
      Width           =   1275
   End
   Begin VB.Label search_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   18
      Top             =   400
      Width           =   600
   End
   Begin VB.Label totalprice_lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "1000.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   69.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   5520
      TabIndex        =   17
      Top             =   5640
      Width           =   4575
   End
   Begin VB.Label stockname_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6720
      TabIndex        =   13
      Top             =   480
      Width           =   45
   End
   Begin VB.Label howmany_lbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1335
      TabIndex        =   9
      Top             =   5760
      Width           =   225
   End
   Begin VB.Label howmany_lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "How many order?"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label numberofstocks_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13321
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6960
      TabIndex        =   7
      Top             =   1560
      Width           =   45
   End
   Begin VB.Label numberofstocks_lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stocks Left:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5520
      TabIndex        =   6
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label price_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13321
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6840
      TabIndex        =   5
      Top             =   1200
      Width           =   45
   End
   Begin VB.Label price_lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5520
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label stocktype_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6840
      TabIndex        =   3
      Top             =   840
      Width           =   45
   End
   Begin VB.Label category_lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5520
      TabIndex        =   2
      Top             =   840
      Width           =   765
   End
   Begin VB.Label stockname_lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Name:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5520
      TabIndex        =   0
      Top             =   480
      Width           =   990
   End
End
Attribute VB_Name = "logging_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub addorder_btn_Click()
If Not DuplicateName Then
    temporaryOrdersTB.AddNew
    temporaryOrdersTB.Fields(0) = Me.stockname_lbl
    temporaryOrdersTB.Fields(1) = Me.price_lbl
    temporaryOrdersTB.Fields(2) = Me.howmany_lbl
    temporaryOrdersTB.Fields(3) = Me.howmany_lbl * Me.price_lbl
    temporaryOrdersTB.Update
Else
    Conn.Execute "UPDATE temporaryOrders SET Quantity=" & Val(temporaryOrdersTB.Fields(2).Value) + Val(Me.howmany_lbl) & ", TotalPrice=" & (Val(Me.howmany_lbl) * Val(Me.price_lbl)) + Val(temporaryOrdersTB.Fields(3).Value) & " WHERE ItemName='" & temporaryOrdersTB.Fields(0).Value & "'"
    temporaryOrdersTB.Requery
End If

Me.totalprice_lbl = Format(GetTotalPrice("TotalPrice", "temporaryOrders"), "0.00")

stockname_txtbox.Text = ""
stockrs.filter = adFilterNone

Me.orders_datgrid.Columns(1).NumberFormat = "0.00"
Me.orders_datgrid.Columns(3).NumberFormat = "0.00"

temporaryOrdersTB.MoveLast

Call RefreshGrid
Call SetStockInfo
Call ProcessOn
Call CheckStock
Me.stockname_txtbox.SetFocus
End Sub

Public Sub cancel_btn_Click()
stockrs.filter = adFilterNone
Unload Me
End Sub

Public Sub decrement_btn_Click()
If howmany_lbl <> 0 Then
    howmany_lbl = howmany_lbl - 1
    Me.numberofstocks_lbl = Me.numberofstocks_lbl + 1
End If
Call ButtonsBehavior(Me)
End Sub

Public Sub delorder_btn_Click()
Call SetWarningForm("Delete Order", "Delete this order?", True, "Delete", True, "Cancel", "DeleteOrder")
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call ShortCutProblems(KeyCode, Shift, Me)
End Sub

Public Sub RefreshGrid()
Me.orders_datgrid.Columns(0).Width = (Me.orders_datgrid.Width / 4) - 140
Me.orders_datgrid.Columns(1).Width = (Me.orders_datgrid.Width / 4) - 140
Me.orders_datgrid.Columns(2).Width = (Me.orders_datgrid.Width / 4) - 140
Me.orders_datgrid.Columns(3).Width = (Me.orders_datgrid.Width / 4) - 140
End Sub

Public Sub Form_Load()
Call SetupLoggingForm
End Sub

Public Sub increment_btn_Click()
If numberofstocks_lbl <> 0 Then
    howmany_lbl = howmany_lbl + 1
    Me.numberofstocks_lbl = Me.numberofstocks_lbl - 1
Else
    Call SetWarningForm("Out Of Stock", "Restock Now", False, "", True, "Ok")
End If
Call ButtonsBehavior(Me)
End Sub

Public Sub process_btn_Click()
Dim StockHold As Integer
Dim MinStocks As String
Dim IsMinStock As Boolean
Dim ItemsSold As String

counter = 0

OrdersTB.AddNew
OrdersTB.Fields(0) = Me.orderno_lbl
OrdersTB.Fields(1) = Me.orderdate_lbl
OrdersTB.Update

Call Tables
    
temporaryOrdersTB.MoveFirst
ItemsSold = temporaryOrdersTB.Fields(0).Value

While counter < temporaryOrdersTB.RecordCount
    If counter <> 0 Then
        ItemsSold = ItemsSold & ", " & temporaryOrdersTB.Fields(0).Value
    End If
    
    OrderItemsTB.AddNew
    OrderItemsTB.Fields(0) = Me.orderno_lbl
    OrderItemsTB.Fields(1) = GetItemID(temporaryOrdersTB.Fields(0).Value)
    OrderItemsTB.Fields(2) = temporaryOrdersTB.Fields(2).Value
    OrderItemsTB.Update
    stockrs.filter = "StockID LIKE '" & GetItemID(temporaryOrdersTB.Fields(0).Value) & "'"
    StockHold = Val(stockrs.Fields(4).Value) - Val(temporaryOrdersTB.Fields(2).Value)
    Conn.Execute "UPDATE Stocks SET StockOnHold=" & StockHold & " WHERE StockID=" & GetItemID(temporaryOrdersTB.Fields(0).Value)
    If StockHold <= stockrs.Fields(6).Value Then
        If Not IsMinStock Then
            MinStocks = stockrs.Fields(1).Value
            IsMinStock = True
        Else
            MinStocks = MinStocks & ", " & stockrs.Fields(1).Value
        End If
    End If
    counter = counter + 1
    temporaryOrdersTB.MoveNext
Wend

orderrs.AddNew
orderrs.Fields(0) = Me.orderno_lbl
orderrs.Fields(1) = Me.orderdate_lbl
orderrs.Fields(2) = ItemsSold
orderrs.Fields(3) = temporaryOrdersTB.RecordCount
orderrs.Fields(4) = Me.totalprice_lbl
orderrs.Update

main_frm.grandtotal_lbl = Format(Val(main_frm.grandtotal_lbl) + Val(Me.totalprice_lbl), "0.00")

stockrs.filter = adFilterNone
stockrs.Requery
Me.totalprice_lbl = "0.00"
Me.orderno_lbl = GetOrderNumber

Call ClearOrders
Call ButtonsBehavior(Me)
Call ButtonsBehavior(main_frm)
Call SetStockInfo
Call RefreshGrid
'Call SetupDataGrids(False, True)
Me.stockname_txtbox.SetFocus

'Call GrandTotalPrice
ungrouprs.Requery
daters.Requery
orderrs.Requery
itemrs.Requery

Call SetupDataGrids(True, True)

Call LogsMoveLast

If IsMinStock Then
    Call SetWarningForm("Minimum Stock Reached", MinStocks & " reached its minimum stock, reorder to restock now.", False, "", True, "Ok")
    MinStocks = ""
    IsMinStock = False
End If
End Sub

Public Sub stockname_txtbox_Change()
If stockname_txtbox.Text = "" Then
    stockrs.filter = adFilterNone
Else
    stockrs.filter = "StockName LIKE '" & stockname_txtbox.Text & "*'"
End If

Call SetupModule.SetupDataGrids(True, False)
Call SetupModule.SetStockInfo
Call ProcessOn
Call CheckOutOfStock
End Sub
