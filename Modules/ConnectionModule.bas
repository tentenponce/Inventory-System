Attribute VB_Name = "ConnectionModule"
Option Explicit
Public Conn, Conn2 As ADODB.Connection
Public stockrs, StocksTB, OrderItemsTB, OrdersTB, CategoriesTB, StockCategoryTB, temporaryOrdersTB, temprs, ExpensesTB, InitialMoneyTB, CashOnHoldTB, SuppliersTB, MinMaxStockTB, SuppProductsTB, TemporarySuppProductsTB, SuppHistoryTB As ADODB.Recordset
Public ungrouprs, orderrs, itemrs, daters, restockrs As ADODB.Recordset
Sub connect()
Set Conn = New ADODB.Connection
    Conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\MainDB.accdb;Jet OLEDB:Database Password=error404;Persist Security Info=False;"
Set Conn2 = New ADODB.Connection
    Conn2.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\LogsDB.accdb;Jet OLEDB:Database Password=error404;Persist Security Info=False;"
Set stockrs = New ADODB.Recordset
    stockrs.ActiveConnection = Conn
    stockrs.CursorLocation = adUseClient
    stockrs.CursorType = adOpenDynamic
    stockrs.LockType = adLockOptimistic
    stockrs.Source = "SELECT * FROM StockInfo"
    stockrs.Open
Set ungrouprs = New ADODB.Recordset
    ungrouprs.ActiveConnection = Conn
    ungrouprs.CursorLocation = adUseClient
    ungrouprs.CursorType = adOpenDynamic
    ungrouprs.LockType = adLockOptimistic
    ungrouprs.Source = "SELECT * FROM LogsPerItem"
    ungrouprs.Open
Set daters = New ADODB.Recordset
    daters.ActiveConnection = Conn
    daters.CursorLocation = adUseClient
    daters.CursorType = adOpenDynamic
    daters.LockType = adLockOptimistic
    daters.Source = "SELECT * FROM DateGroup"
    daters.Open
Set itemrs = New ADODB.Recordset
    itemrs.ActiveConnection = Conn
    itemrs.CursorLocation = adUseClient
    itemrs.CursorType = adOpenDynamic
    itemrs.LockType = adLockOptimistic
    itemrs.Source = "SELECT * FROM StockGroup"
    itemrs.Open
Set orderrs = New ADODB.Recordset
    orderrs.ActiveConnection = Conn2
    orderrs.CursorLocation = adUseClient
    orderrs.CursorType = adOpenDynamic
    orderrs.LockType = adLockOptimistic
    orderrs.Source = "SELECT * FROM Logs"
    orderrs.Open
Set restockrs = New ADODB.Recordset
    restockrs.ActiveConnection = Conn
    restockrs.CursorLocation = adUseClient
    restockrs.CursorType = adOpenDynamic
    restockrs.LockType = adLockOptimistic
    restockrs.Source = "SELECT * FROM Restock"
    restockrs.Open
Set temprs = New ADODB.Recordset
    temprs.ActiveConnection = Conn
    temprs.CursorLocation = adUseClient
    temprs.CursorType = adOpenDynamic
    temprs.LockType = adLockOptimistic
Call Tables
End Sub
Sub Tables()
Set StocksTB = New ADODB.Recordset
    StocksTB.CursorLocation = adUseClient
    StocksTB.Open "SELECT * FROM Stocks", Conn, adOpenDynamic, adLockOptimistic
Set OrderItemsTB = New ADODB.Recordset
    OrderItemsTB.CursorLocation = adUseClient
    OrderItemsTB.Open "SELECT * FROM OrderItems", Conn, adOpenDynamic, adLockOptimistic
Set OrdersTB = New ADODB.Recordset
    OrdersTB.CursorLocation = adUseClient
    OrdersTB.Open "SELECT * FROM Orders", Conn, adOpenDynamic, adLockOptimistic
Set StockCategoryTB = New ADODB.Recordset
    StockCategoryTB.CursorLocation = adUseClient
    StockCategoryTB.Open "SELECT * FROM StockCategory", Conn, adOpenDynamic, adLockOptimistic
Set CategoriesTB = New ADODB.Recordset
    CategoriesTB.CursorLocation = adUseClient
    CategoriesTB.Open "SELECT * FROM Categories", Conn, adOpenDynamic, adLockOptimistic
Set temporaryOrdersTB = New ADODB.Recordset
    temporaryOrdersTB.CursorLocation = adUseClient
    temporaryOrdersTB.Open "SELECT * FROM temporaryOrders", Conn, adOpenDynamic, adLockOptimistic
Set ExpensesTB = New ADODB.Recordset
    ExpensesTB.CursorLocation = adUseClient
    ExpensesTB.Open "SELECT * FROM Expenses", Conn, adOpenDynamic, adLockOptimistic
Set InitialMoneyTB = New ADODB.Recordset
    InitialMoneyTB.CursorLocation = adUseClient
    InitialMoneyTB.Open "SELECT * FROM InitialMoney", Conn, adOpenDynamic, adLockOptimistic
Set CashOnHoldTB = New ADODB.Recordset
    CashOnHoldTB.CursorLocation = adUseClient
    CashOnHoldTB.Open "SELECT * FROM CashOnHold", Conn, adOpenDynamic, adLockOptimistic
Set SuppliersTB = New ADODB.Recordset
    SuppliersTB.CursorLocation = adUseClient
    SuppliersTB.Open "SELECT * FROM Suppliers", Conn, adOpenDynamic, adLockOptimistic
Set MinMaxStockTB = New ADODB.Recordset
    MinMaxStockTB.CursorLocation = adUseClient
    MinMaxStockTB.Open "SELECT * FROM MinMaxStock", Conn, adOpenDynamic, adLockOptimistic
Set SuppProductsTB = New ADODB.Recordset
    SuppProductsTB.CursorLocation = adUseClient
    SuppProductsTB.Open "SELECT * FROM SupplierProducts", Conn, adOpenDynamic, adLockOptimistic
Set SuppliersTB = New ADODB.Recordset
    SuppliersTB.CursorLocation = adUseClient
    SuppliersTB.Open "SELECT * FROM Suppliers", Conn, adOpenDynamic, adLockOptimistic
Set TemporarySuppProductsTB = New ADODB.Recordset
    TemporarySuppProductsTB.CursorLocation = adUseClient
    TemporarySuppProductsTB.Open "SELECT * FROM TemporarySuppProducts", Conn, adOpenDynamic, adLockOptimistic
Set SuppHistoryTB = New ADODB.Recordset
    SuppHistoryTB.CursorLocation = adUseClient
    SuppHistoryTB.Open "SELECT * FROM SupplyHistory", Conn, adOpenDynamic, adLockOptimistic
End Sub

Sub Main()
connect
End Sub
