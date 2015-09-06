Attribute VB_Name = "ReportModule"
Public Sub Filtering(fromdate As Date, todate As Date)
titlebar_frm.todate_lbl = Format(todate, "dd/mm/yyyy")
titlebar_frm.fromdate_lbl = Format(fromdate, "dd/mm/yyyy")
ReportFilter = "DateOrdered >= #" & titlebar_frm.fromdate_lbl & "# AND DateOrdered <= #" & titlebar_frm.todate_lbl & "#"

ungrouprs.filter = ReportFilter
daters.filter = ReportFilter
orderrs.filter = ReportFilter

ExpensesTB.filter = "ExpenseDate >= #" & titlebar_frm.fromdate_lbl & "# AND ExpenseDate <= #" & titlebar_frm.todate_lbl & "#"
InitialMoneyTB.filter = "LogDate >= #" & titlebar_frm.fromdate_lbl & "# AND LogDate <= #" & titlebar_frm.todate_lbl & "#"
CashOnHoldTB.filter = "LogDate >= #" & titlebar_frm.fromdate_lbl & "# AND LogDate <= #" & titlebar_frm.todate_lbl & "#"

Call GrandTotalPrice
Call CalculateCashOnHold
Call CalculateInitialMoney
End Sub

Public Sub GrandTotalPrice()
ungrouprs.Requery
daters.Requery
orderrs.Requery
itemrs.Requery
main_frm.grandtotal_lbl = Format(GetTotalPrice("TotalPrice", "LogsPerItem", " WHERE " & "DateOrdered >= #" & Format(titlebar_frm.fromdate_lbl, "mm/dd/yyyy") & "# AND DateOrdered <= #" & Format(titlebar_frm.todate_lbl, "mm/dd/yyyy") & "#"), "0.00")
Call SetupDataGrids(False, True)
End Sub

Public Sub CashOnHold()
Dim grandtotal As Long
grandtotal = 0
counter = 0

If main_frm.grandtotal_lbl <> "" Then
    expenses_frm.grandtotal_lbl = main_frm.grandtotal_lbl
Else
    expenses_frm.grandtotal_lbl = "0"
End If

If ExpensesTB.RecordCount <> 0 Then
    ExpensesTB.MoveFirst
    While counter < ExpensesTB.RecordCount
        grandtotal = Val(grandtotal) + Val(ExpensesTB.Fields(2).Value)
        counter = counter + 1
        ExpensesTB.MoveNext
    Wend
    expenses_frm.exp_lbl = Format(grandtotal, "0.00")
    grandtotal = 0
    ExpensesTB.MoveLast
Else
    grandtotal = 0
    expenses_frm.exp_lbl = Format(grandtotal, "0.00")
End If

expenses_frm.cash_lbl = Format((expenses_frm.grandtotal_lbl - expenses_frm.exp_lbl) + expenses_frm.initial_lbl, "0.00")
End Sub

Public Sub CalculateCashOnHold()
counter = 0
expenses_frm.cashonhold_lbl = "0.00"

If CashOnHoldTB.RecordCount <> 0 Then
    CashOnHoldTB.MoveFirst
    While counter < CashOnHoldTB.RecordCount
        expenses_frm.cashonhold_lbl = Format(Val(expenses_frm.cashonhold_lbl) + CashOnHoldTB.Fields(0).Value, "0.00")
        counter = counter + 1
        CashOnHoldTB.MoveNext
    Wend
Else
    expenses_frm.cashonhold_lbl = "0.00"
    expenses_frm.cashonhold_lbl.Refresh
End If
End Sub

Public Sub CalculateInitialMoney()
counter = 0
expenses_frm.initial_lbl = "0.00"

If InitialMoneyTB.RecordCount <> 0 Then
    InitialMoneyTB.MoveFirst
    While counter < InitialMoneyTB.RecordCount
        expenses_frm.initial_lbl = Format(Val(expenses_frm.initial_lbl) + InitialMoneyTB.Fields(0).Value, "0.00")
        counter = counter + 1
        InitialMoneyTB.MoveNext
    Wend
Else
    expenses_frm.initial_lbl = "0.00"
    expenses_frm.initial_lbl.Refresh
End If

Call CashOnHold
End Sub

Public Sub GetTopSoldItem()
counter2 = 0
temprs.Open "SELECT StockName, COUNT(OrderNo) FROM LogsPerItem WHERE " & ReportFilter & " GROUP BY StockName ORDER BY COUNT(OrderNo) DESC"
While counter2 < 5
    If Not temprs.EOF And Not temprs.BOF Then
        Do While IsUnfiltered(GetItemCategory(temprs.Fields(0).Value))
            temprs.MoveNext
            If temprs.BOF Or temprs.EOF Then
                Exit Do
            End If
        Loop
        If Not temprs.EOF And Not temprs.BOF Then
            pos_frm.topsold_lbl(counter2) = temprs.Fields(0).Value
            temprs.MoveNext
        End If
    Else
        pos_frm.topsold_lbl(counter2) = "N/A"
    End If
    counter2 = counter2 + 1
Wend
temprs.Close
End Sub

Public Sub GetTopNumSoldItem()
counter2 = 0
temprs.Open "SELECT StockName, SUM(Quantity) FROM LogsPerItem WHERE " & ReportFilter & " GROUP BY StockName ORDER BY SUM(Quantity) DESC"
While counter2 < 5
    If Not temprs.EOF And Not temprs.BOF Then
        Do While IsUnfiltered(GetItemCategory(temprs.Fields(0).Value))
            temprs.MoveNext
            If temprs.BOF Or temprs.EOF Then
                Exit Do
            End If
        Loop
        If Not temprs.BOF And Not temprs.EOF Then
            pos_frm.topnumsold_lbl(counter2) = temprs.Fields(0).Value
            pos_frm.topnumsold_lbl(counter2 + 5) = "x " & temprs.Fields(1).Value
            temprs.MoveNext
        End If
    Else
        pos_frm.topnumsold_lbl(counter2) = "N/A"
        pos_frm.topnumsold_lbl(counter2 + 5) = "N/A"
    End If
    counter2 = counter2 + 1
Wend
temprs.Close
End Sub

Public Function GetAverageIncome(Optional exp As Double) As Double
If ungrouprs.RecordCount <> 0 Then
    counter = 0
    
    temprs.Open "SELECT DISTINCT(DateOrdered) FROM LogsPerItem WHERE " & ReportFilter
    While counter < temprs.RecordCount
        counter = counter + 1
    Wend
    temprs.Close
    GetAverageIncome = (main_frm.grandtotal_lbl - exp) / counter
Else
    GetAverageIncome = 0
End If
End Function

Public Function GetAverageExpense() As Double
If ExpensesTB.RecordCount <> 0 Then
    counter = 0
    
    temprs.Open "SELECT DISTINCT(ExpenseDate) FROM Expenses WHERE " & ExpensesTB.filter
    While counter < ExpensesTB.RecordCount
        counter = counter + 1
    Wend
    temprs.Close
    GetAverageExpense = pos_frm.totalexp_lbl / counter
Else
    GetAverageExpense = 0
End If
End Function

Public Function IsUnfiltered(ItemCategory As String) As Boolean
counter = 0
IsUnfiltered = False

While counter < pos_frm.unfiltered_cmbox.ListCount
    If pos_frm.unfiltered_cmbox.List(counter) = ItemCategory Or ItemCategory = "N/A" Then
        IsUnfiltered = True
        Exit Function
    End If
    counter = counter + 1
Wend
End Function
