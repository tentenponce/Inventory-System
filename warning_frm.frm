VERSION 5.00
Begin VB.Form warning_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   Icon            =   "warning_frm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton ok_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Ok"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label desc_lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   3345
   End
   Begin VB.Label head_lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Head Description"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "warning_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Restored As Boolean

Public Sub cancel_btn_Click()
Unload Me
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Call cancel_btn_Click
End If
End Sub

Public Sub Form_Load()
Call SetupShadow.SetupForm(Me)
End Sub

Public Sub ok_btn_Click()
If Me.Caption = "DeleteStock" Then
    Conn.Execute "DELETE * FROM Stocks WHERE StockID = " & stockrs.Fields(0).Value
    stockrs.filter = adFilterNone
    stockrs.Requery
    
    Call SetupDataGrids(True, False)
    Call SetupModule.SetStockInfo
    Unload Me
ElseIf Me.Caption = "DeleteOrder" Then
    Conn.Execute "DELETE * FROM temporaryOrders WHERE ItemName='" & temporaryOrdersTB.Fields(0).Value & "'"
    temporaryOrdersTB.Requery
    logging_frm.totalprice_lbl = Format(GetTotalPrice("TotalPrice", "temporaryOrders"), "0.00")
    If temporaryOrdersTB.RecordCount <> 0 Then
        temporaryOrdersTB.MoveLast
    End If
    Call logging_frm.RefreshGrid
    Call SetStockInfo
    Call ProcessOn
    Call CheckStock
    Call ButtonsBehavior(main_frm)
    logging_frm.stockname_txtbox.SetFocus
    Unload Me
ElseIf Me.Caption = "DeleteRec" Then
    Call ReturnStock
    If titlebar_frm.groupby_cmbox.Text = "Ungroup" Then
        Conn.Execute "DELETE * FROM OrderItems WHERE OrderNo=" & ungrouprs.Fields(0).Value & " AND StockID=" & GetItemID(ungrouprs.Fields(2).Value)
    Else
        Conn.Execute "DELETE * FROM OrderItems WHERE OrderNo=" & orderrs.Fields(0).Value
        Conn2.Execute "DELETE * FROM Logs WHERE OrderNo=" & orderrs.Fields(0).Value
        Conn.Execute "DELETE * FROM Orders WHERE OrderNo=" & orderrs.Fields(0).Value
    End If
    Call GrandTotalPrice
    LogsMoveLast
    Call ButtonsBehavior(main_frm)
    Unload Me
ElseIf Me.Caption = "DeleteSupp" Then
    Conn.Execute "DELETE * FROM Suppliers WHERE SupplierID=" & GetSuppID(restock_frm.supp_cmbox.Text)
    SuppliersTB.Requery
    Call SetSuppliers
    Call SetRestockProducts
    Unload Me
ElseIf Me.Caption = "DeleteExp" Then
    ExpensesTB.Delete
    ExpensesTB.Requery
    Call ButtonsBehavior(main_frm)
    Call expenses_frm.RefreshGrid
    If ExpensesTB.RecordCount <> 0 Then
        ExpensesTB.MoveLast
    End If
    Unload Me
ElseIf Me.Caption = "Restock" Then
    SuppHistoryTB.AddNew
    SuppHistoryTB.Fields(0) = DateValue(Now)
    SuppHistoryTB.Fields(1) = GetSuppID(restock_frm.supp_cmbox.Text)
    SuppHistoryTB.Fields(2) = GetItemID(restockrs.Fields(1).Value)
    SuppHistoryTB.Fields(3) = restock_frm.howmany_lbl
    SuppHistoryTB.Update
    SuppHistoryTB.Requery
    
    stockrs.filter = "StockID LIKE " & GetItemID(restockrs.Fields(1).Value)
    Conn.Execute "UPDATE Stocks SET StockOnHold=" & Val(main_frm.stockonhold_txtbox.Text) + Val(restock_frm.howmany_lbl) & " WHERE StockID=" & GetItemID(restockrs.Fields(1).Value)
    stockrs.filter = adFilterNone
    stockrs.Requery
    Call SetupModule.SetStockInfo
    Call SetupModule.SetupDataGrids(True, False)
    Call ButtonsBehavior(main_frm)
    restock_frm.howmany_lbl = "0"
    Call SetRestockProducts
    Unload Me
ElseIf Me.Caption = "ExitSystem" Then
    Timers.statusbar.Enabled = False
    Timers.mainfrm_close.Enabled = True
    Me.Hide
End If
End Sub
