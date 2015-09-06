Attribute VB_Name = "FunctionsModule"
Public Function GetCategoryCount()
GetCategoryCount = CategoriesTB.RecordCount
End Function

Public Function GetCategoryID(CategoryName As String)
CategoriesTB.Requery
CategoriesTB.filter = "Category LIKE '" & CategoryName & "'"
GetCategoryID = CategoriesTB.Fields(0).Value
CategoriesTB.filter = adFilterNone
End Function

Public Function GetItemID(ItemName As String)
StocksTB.Requery
StocksTB.filter = "StockName LIKE '" & ItemName & "'"
GetItemID = StocksTB.Fields(0).Value
StocksTB.filter = adFilterNone
End Function

Public Function GetSuppID(SuppName As String)
SuppliersTB.Requery
SuppliersTB.filter = "SupplierName LIKE '" & SuppName & "'"
If SuppliersTB.RecordCount <> 0 Then
    GetSuppID = Format(SuppliersTB.Fields(0).Value, "000")
Else
    GetSuppID = ""
End If
SuppliersTB.filter = adFilterNone
End Function

Public Function GetTotalPrice(columnName As String, TableName As String, Optional filter As String) As Double
temprs.Open "SELECT SUM(" & columnName & ") FROM " & TableName & filter
If temprs.RecordCount <> 0 And temprs.Fields(0) <> "" Then
    GetTotalPrice = Format(temprs.Fields(0), "0.00")
End If
temprs.Close
End Function

Public Function GetOrderNumber()
temprs.Open "SELECT MAX(OrderNo) FROM Orders"

If OrdersTB.RecordCount <> 0 Then
    GetOrderNumber = temprs.Fields(0) + 1
Else
    GetOrderNumber = 1
End If
temprs.Close
End Function

Public Function DuplicateName() As Boolean
counter = 0

If temporaryOrdersTB.RecordCount <> 0 Then
    temporaryOrdersTB.MoveFirst
    While counter < temporaryOrdersTB.RecordCount
        If logging_frm.stockname_lbl = temporaryOrdersTB.Fields(0).Value Then
            DuplicateName = True
            Exit Function
        End If
        counter = counter + 1
        temporaryOrdersTB.MoveNext
    Wend
End If
End Function

Public Function GetMonthNumber(month_cmbox As String)
Select Case month_cmbox
    Case "January"
        GetMonthNumber = "01"
    Case "February"
        GetMonthNumber = "02"
    Case "March"
        GetMonthNumber = "03"
    Case "April"
        GetMonthNumber = "04"
    Case "May"
        GetMonthNumber = "05"
    Case "June"
        GetMonthNumber = "06"
    Case "July"
        GetMonthNumber = "07"
    Case "August"
        GetMonthNumber = "08"
    Case "September"
        GetMonthNumber = "09"
    Case "October"
        GetMonthNumber = "10"
    Case "November"
        GetMonthNumber = "11"
    Case "December"
        GetMonthNumber = "12"
End Select
End Function

Public Function GetDate(monthh As String, dayy As String, yearr As String) As Date
GetDate = monthh & "/" & dayy & "/" & yearr
End Function

Public Function GetItemCategory(ItemName As String) As String
Dim OrigFilter As String
OrigFilter = stockrs.filter
stockrs.filter = "StockID LIKE " & GetItemID(ItemName)
If stockrs.RecordCount <> 0 Then
    GetItemCategory = stockrs.Fields(2).Value
Else
    GetItemCategory = "N/A"
End If
If OrigFilter = "0" Then
    stockrs.filter = adFilterNone
Else
    stockrs.filter = OrigFilter
End If
End Function
