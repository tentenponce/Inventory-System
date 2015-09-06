Attribute VB_Name = "KeyCodeModule"
Public Sub ShortCutProblems(WhatKey As Integer, Cheft As Integer, formname As Form)
If WhatKey = vbKeyEscape Then
    If formname.Name = "main_frm" Then
        If EditMode Or AddMode Then
            If main_frm.cancel_btn.Enabled Then
                Call main_frm.cancel_btn_Click
            End If
        Else
            Call SetWarningForm("Exit System", "Exit System? (All datas are saved.)", True, "Exit", True, "Cancel", "ExitSystem")
        End If
    ElseIf formname.Name = "expenses_frm" Then
        If expenses_frm.cancel_btn.Enabled Then
            Call expenses_frm.cancel_btn_Click
        End If
    Else
        Unload formname
    End If
ElseIf WhatKey = 13 Then
    If formname.Name = "logging_frm" Then
        If logging_frm.addorder_btn.Enabled Then
            logging_frm.addorder_btn_Click
        End If
    ElseIf formname.Name = "main_frm" Then
        If main_frm.ok_btn.Enabled Then
            main_frm.ok_btn_Click
        End If
    End If
ElseIf WhatKey = vbKeyF2 Then
    If formname.Name = "logging_frm" Then
        If logging_frm.process_btn.Enabled Then
            logging_frm.process_btn_Click
        End If
    ElseIf formname.Name = "main_frm" Then
        If main_frm.addrec_btn.Enabled Then
            main_frm.addrec_btn_Click
        End If
    ElseIf formname.Name = "restock_frm" Then
        If restock_frm.restock_btn.Enabled Then
            Call restock_frm.restock_btn_Click
        End If
    End If
ElseIf WhatKey = vbKeyF3 Then
    If formname.Name = "main_frm" Then
        If main_frm.expense_btn.Enabled Then
            Call main_frm.expense_btn_Click
        End If
    End If
ElseIf WhatKey = vbKeyF5 Then
    If formname.Name = "main_frm" Then
        If main_frm.refreshstock_btn.Enabled Then
            Call main_frm.refreshstock_btn_Click
        End If
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyUp Then
    If stockrs.RecordCount <> 0 Then
        If formname.Name = "logging_frm" Then
            If stockrs.AbsolutePosition = 1 Then
                stockrs.MoveLast
            ElseIf stockrs.RecordCount <> 0 Then
                stockrs.MovePrevious
            End If
        ElseIf formname.Name = "main_frm" Then
            If Not EditMode And Not AddMode Then
                If stockrs.AbsolutePosition = 1 Then
                    stockrs.MoveLast
                ElseIf stockrs.RecordCount <> 0 Then
                    stockrs.MovePrevious
                End If
            End If
        ElseIf formname.Name = "restock_frm" Then
            If restock_frm.supp_cmbox.Enabled Then
                If restockrs.AbsolutePosition = 1 Then
                    restockrs.MoveLast
                ElseIf restockrs.RecordCount <> 0 Then
                    restockrs.MovePrevious
                End If
                restock_frm.howmany_lbl = "0"
                restock_frm.restock_btn.SetFocus
            Else
                If stockrs.AbsolutePosition = 1 Then
                    stockrs.MoveLast
                ElseIf stockrs.RecordCount <> 0 Then
                    stockrs.MovePrevious
                End If
            End If
        End If
        
        Call CheckOutOfStock
        Call SetStockInfo
        
        If formname.Name = "logging_frm" Then
            Call ProcessOn
            Call CheckStock
        End If
        
        If formname.Name = "restock_frm" Then
            If Not restock_frm.supp_cmbox.Enabled Then
                Call restock_frm.SetStockInfos
            End If
        End If
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyDown Then
    If stockrs.RecordCount <> 0 Then
        If formname.Name = "logging_frm" Then
            If stockrs.AbsolutePosition = stockrs.RecordCount Then
                stockrs.MoveFirst
            ElseIf stockrs.RecordCount <> 0 Then
                stockrs.MoveNext
            End If
            Call CheckStock
        ElseIf formname.Name = "main_frm" Then
            If Not EditMode And Not AddMode Then
                If stockrs.AbsolutePosition = stockrs.RecordCount Then
                    stockrs.MoveFirst
                ElseIf stockrs.RecordCount <> 0 Then
                    stockrs.MoveNext
                End If
            End If
        ElseIf formname.Name = "restock_frm" Then
            If restock_frm.supp_cmbox.Enabled Then
                If restockrs.AbsolutePosition = restockrs.RecordCount Then
                    restockrs.MoveFirst
                ElseIf stockrs.RecordCount <> 0 Then
                    restockrs.MoveNext
                End If
                restock_frm.howmany_lbl = "0"
                restock_frm.restock_btn.SetFocus
            Else
                If stockrs.AbsolutePosition = stockrs.RecordCount Then
                    stockrs.MoveFirst
                ElseIf stockrs.RecordCount <> 0 Then
                    stockrs.MoveNext
            End If
            End If
        End If
        
        Call CheckOutOfStock
        Call SetStockInfo
        
        If formname.Name = "logging_frm" Then
            Call ProcessOn
            Call CheckStock
        End If
        
        If formname.Name = "restock_frm" Then
            If Not restock_frm.supp_cmbox.Enabled Then
                Call restock_frm.SetStockInfos
            End If
        End If
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyA Then
    If formname.Name = "main_frm" Then
        If main_frm.addstock_btn.Enabled Then
            Call main_frm.addstock_btn_Click
        End If
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyR Then
    If formname.Name = "main_frm" Then
        If main_frm.report_btn.Enabled Then
            Call main_frm.report_btn_Click
        End If
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyH Then
    If formname.Name = "main_frm" Then
        If titlebar_frm.help_btn.Enabled Then
            Call titlebar_frm.help_btn_Click
        End If
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyD Then
    If formname.Name = "main_frm" Then
        If main_frm.delrec_btn.Enabled Then
            Call main_frm.delrec_btn_Click
        End If
    ElseIf formname.Name = "logging_frm" Then
        If logging_frm.delorder_btn.Enabled Then
            logging_frm.delorder_btn_Click
        End If
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyS Then
    If formname.Name = "main_frm" Then
        If titlebar_frm.search_txtbox.Enabled Then
            titlebar_frm.search_txtbox.Text = ""
            titlebar_frm.search_txtbox.SetFocus
        End If
    ElseIf formname.Name = "logging_frm" Then
        logging_frm.stockname_txtbox.Text = ""
        logging_frm.stockname_txtbox.SetFocus
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyE Then
    If formname.Name = "main_frm" Then
        If main_frm.editstock_btn.Enabled Then
            main_frm.editstock_btn_Click
        End If
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyF Then
    If formname.Name = "main_frm" Then
        If titlebar_frm.datefilter_btn.Enabled Then
            titlebar_frm.datefilter_btn_Click
        End If
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyT Then
    If formname.Name = "main_frm" Then
        If main_frm.todaylog_btn.Enabled Then
            main_frm.todaylog_btn_Click
        End If
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyP Then
    If formname.Name = "main_frm" Then
        If main_frm.prevday_btn.Enabled Then
            main_frm.prevday_btn_Click
        End If
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyN Then
    If formname.Name = "main_frm" Then
        If main_frm.nextday_btn.Enabled Then
            main_frm.nextday_btn_Click
        End If
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyL Then
    If formname.Name = "main_frm" Then
        If main_frm.alllog_btn.Enabled Then
            main_frm.alllog_btn_Click
        End If
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyLeft Then
    If formname.Name = "logging_frm" Then
        Call logging_frm.decrement_btn_Click
    ElseIf formname.Name = "restock_frm" Then
        If restock_frm.decrement_btn.Enabled Then
            Call restock_frm.decrement_btn_Click
        End If
    ElseIf formname.Name = "pos_frm" Then
        If pos_frm.tofiltered_btn.Enabled Then
            Call pos_frm.tofiltered_btn_Click
        End If
    End If
ElseIf (Cheft And vbAltMask) = vbAltMask And WhatKey = vbKeyRight Then
    If formname.Name = "logging_frm" Then
        Call logging_frm.increment_btn_Click
    ElseIf formname.Name = "restock_frm" Then
        If restock_frm.increment_btn.Enabled Then
            Call restock_frm.increment_btn_Click
        End If
    ElseIf formname.Name = "pos_frm" Then
        If pos_frm.tounfiltered_btn.Enabled Then
            Call pos_frm.tounfiltered_btn_Click
        End If
    End If
End If
End Sub
