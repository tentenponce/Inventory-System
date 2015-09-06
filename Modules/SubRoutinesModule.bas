Attribute VB_Name = "SubRoutinesModule"
Public Sub CheckOutOfStock()
If stockrs.RecordCount <> 0 Then
    If stockrs.Fields(4).Value <= stockrs.Fields(6).Value Then
        main_frm.stocks_datgrid.SelBookmarks.Add main_frm.stocks_datgrid.Bookmark
    End If
    
    If stockrs.Fields(5).Value < DateValue(Now) Then
        main_frm.stocks_datgrid.SelBookmarks.Add main_frm.stocks_datgrid.Bookmark
    End If
End If
End Sub

Public Sub CheckStock() 'Logging Form
counter = 0

If temporaryOrdersTB.RecordCount <> 0 And logging_frm.stockname_lbl <> "" Then
    temporaryOrdersTB.MoveFirst
    While counter < temporaryOrdersTB.RecordCount
        If logging_frm.stockname_lbl = temporaryOrdersTB.Fields(0).Value Then
            logging_frm.numberofstocks_lbl = Val(logging_frm.numberofstocks_lbl) - Val(temporaryOrdersTB.Fields(2).Value)
        End If
        counter = counter + 1
        temporaryOrdersTB.MoveNext
    Wend
    temporaryOrdersTB.MoveLast
End If
End Sub

Public Sub ClearOrders() 'Logging Form
Conn.Execute "DELETE * FROM temporaryOrders"
Set logging_frm.orders_datgrid.DataSource = temporaryOrdersTB
temporaryOrdersTB.Requery
End Sub

Public Sub LogsMoveLast()
If ungrouprs.RecordCount <> 0 Then
    ungrouprs.MoveLast
    orderrs.MoveLast
    itemrs.MoveLast
    daters.MoveLast
End If
End Sub

Public Sub ReturnStock()
If titlebar_frm.groupby_cmbox.Text = "Ungroup" Then
    stockrs.filter = "StockID LIKE '" & GetItemID(ungrouprs.Fields(2).Value) & "'"
    Call SetStockInfo
    Conn.Execute "UPDATE Stocks SET StockOnHold=" & Val(ungrouprs.Fields(3).Value) + Val(main_frm.stockonhold_txtbox.Text) & " WHERE StockID=" & GetItemID(ungrouprs.Fields(2).Value)
Else
    ungrouprs.filter = "OrderNo LIKE '" & orderrs.Fields(0).Value & "'"
    counter = 0
    ungrouprs.MoveFirst
    While counter < ungrouprs.RecordCount
        stockrs.filter = "StockID LIKE '" & GetItemID(ungrouprs.Fields(2).Value) & "'"
        Call SetStockInfo
        Conn.Execute "UPDATE Stocks SET StockOnHold=" & Val(ungrouprs.Fields(3).Value) + Val(main_frm.stockonhold_txtbox.Text) & " WHERE StockID=" & GetItemID(ungrouprs.Fields(2).Value)
        ungrouprs.MoveNext
        counter = counter + 1
    Wend
    ungrouprs.filter = ReportFilter
End If
stockrs.filter = adFilterNone
stockrs.Requery
Call SetStockInfo
Call SetupDataGrids(True, False)
End Sub

Public Sub MathProblems(txtbox As TextBox, lbl As Label, num As Integer)
If IsNumeric(Val(txtbox.Text)) Then
    lbl = Format(Val(txtbox.Text) * num, "0.00")
    cashonhold_frm.total_lbl = Format(Val(cashonhold_frm.aooo_lbl) + Val(cashonhold_frm.eoo_lbl) + Val(cashonhold_frm.boo_lbl) + Val(cashonhold_frm.aoo_lbl) + Val(cashonhold_frm.eo_lbl) + Val(cashonhold_frm.bo_lbl) + Val(cashonhold_frm.ao_lbl) + Val(cashonhold_frm.e_lbl) + Val(cashonhold_frm.a_lbl), "0.00")
Else
Call SetWarningForm("Content Mismathed", "Pieces must be numbers only.", False, "", True, "Ok")
txtbox.Text = "0"
End If
End Sub

Public Sub ClearSuppEdit()
restock_frm.suppname_txtbox.Text = ""
Conn.Execute "DELETE * FROM TemporarySuppProducts"
TemporarySuppProductsTB.Requery
End Sub

Public Sub SetText(textfile As String, lblbool As Boolean, Optional lbl As Label, Optional lst As ListBox)
Dim sFileText As String
Dim iFileNo As Integer
counter = 0
iFileNo = FreeFile
Open App.Path & "\textsource\" & textfile & ".txt" For Input As #iFileNo
Do While Not EOF(iFileNo)
    Input #iFileNo, sFileText
    If Not lblbool Then
        lst.AddItem sFileText
    Else
        lbl = lbl & sFileText & vbLf
    End If
    counter = counter + 1
Loop
Close #iFileNo
End Sub

Public Sub RandomTipsss()
Dim sFileText As String
Dim iFileNo As Integer
counter = 0
iFileNo = FreeFile
Open App.Path & "\textsource\randomtips.txt" For Input As #iFileNo
Do While Not EOF(iFileNo)
    Input #iFileNo, sFileText
    Randomtipss(counter) = sFileText
    counter = counter + 1
Loop
Randomize
loading_frm.tips_lbl = Randomtipss(Int(Rnd() * 20))
Close #iFileNo
Timers.randomtips.Enabled = True
End Sub
