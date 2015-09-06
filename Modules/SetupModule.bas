Attribute VB_Name = "SetupModule"
Dim lR As Long
Public Sub SetupMainForm()
'Connections/Others
Call SetCategories(main_frm.categ_txtbox)

Set main_frm.stocks_datgrid.DataSource = stockrs
Set main_frm.logs_datgrid.DataSource = ungrouprs

'Left Grid
main_frm.stocks_datgrid.Left = 120
main_frm.stocks_datgrid.Top = titlebar_frm.Height + 120
main_frm.stocks_datgrid.Height = StockGridHeight
main_frm.stocks_datgrid.Width = StockGridWidth

'Stock Components Layout Add/Delete
main_frm.addstock_btn.Top = main_frm.stocks_datgrid.Top + main_frm.stocks_datgrid.Height + Gap
main_frm.addstock_btn.Left = 120
main_frm.addstock_btn.Height = SmallButtonsHeight
main_frm.addstock_btn.Width = (StockGridWidth / 2) - (Gap / 2) '*0.25
main_frm.delstock_btn.Top = main_frm.addstock_btn.Top
'main_frm.delstock_btn.Left = 120 + main_frm.addstock_btn.Width + Gap
main_frm.delstock_btn.Left = 120
main_frm.delstock_btn.Height = SmallButtonsHeight
main_frm.delstock_btn.Width = main_frm.addstock_btn.Width
main_frm.restock_btn.Top = main_frm.addstock_btn.Top
main_frm.restock_btn.Left = main_frm.delstock_btn.Left + main_frm.delstock_btn.Width + Gap
main_frm.restock_btn.Height = SmallButtonsHeight
'main_frm.restock_btn.Width = StockGridWidth - main_frm.addstock_btn.Width - main_frm.delstock_btn.Width - (Gap * 2)
'main_frm.restock_btn.Width = StockGridWidth - (main_frm.addstock_btn.Width / 2) - main_frm.delstock_btn.Width - (Gap * 2)
main_frm.restock_btn.Width = main_frm.addstock_btn.Width
main_frm.outofstock_btn.Top = main_frm.addstock_btn.Top + main_frm.addstock_btn.Height + Gap
main_frm.outofstock_btn.Left = main_frm.addstock_btn.Left
main_frm.outofstock_btn.Height = SmallButtonsHeight
main_frm.outofstock_btn.Width = (StockGridWidth / 2) - (Gap / 2)
main_frm.refreshstock_btn.Top = main_frm.outofstock_btn.Top
main_frm.refreshstock_btn.Left = main_frm.outofstock_btn.Left + main_frm.outofstock_btn.Width + Gap
main_frm.refreshstock_btn.Height = SmallButtonsHeight
main_frm.refreshstock_btn.Width = main_frm.outofstock_btn.Width

main_frm.editstock_btn.Top = Screen.Height - SmallButtonsHeight - StatusBarHeight - (Gap * 2)
main_frm.editstock_btn.Left = 120 * 2
main_frm.editstock_btn.Height = SmallButtonsHeight
main_frm.editstock_btn.Width = StockGridWidth * 0.4
main_frm.ok_btn.Left = main_frm.editstock_btn.Left + main_frm.editstock_btn.Width + Gap
main_frm.ok_btn.Top = main_frm.editstock_btn.Top
main_frm.ok_btn.Height = SmallButtonsHeight
main_frm.ok_btn.Width = StockGridWidth * 0.25
main_frm.cancel_btn.Left = main_frm.ok_btn.Left + main_frm.ok_btn.Width + Gap
main_frm.cancel_btn.Top = main_frm.editstock_btn.Top
main_frm.cancel_btn.Height = SmallButtonsHeight
main_frm.cancel_btn.Width = StockGridWidth * 0.25

'Edit Components Layout
main_frm.stockid_lbl.Left = 120 * 2
main_frm.stockid_lbl.Top = Screen.Height - (SmallButtonsHeight + 240) - StatusBarHeight - ((main_frm.stockname_txtbox.Height + 120) * 5.7) - 120
main_frm.stockname_lbl.Left = 120 * 2
main_frm.stockname_lbl.Top = main_frm.stockid_lbl.Top + main_frm.stockid_lbl.Height + Gap + 60
main_frm.categ_lbl.Left = 120 * 2
main_frm.categ_lbl.Top = main_frm.stockname_lbl.Top + main_frm.stockname_lbl.Height + Gap + 60
main_frm.price_lbl.Left = 120 * 2
main_frm.price_lbl.Top = main_frm.categ_lbl.Top + main_frm.categ_lbl.Height + Gap + 60
main_frm.stockonhold_lbl.Left = 120 * 2
main_frm.stockonhold_lbl.Top = main_frm.price_lbl.Top + main_frm.price_lbl.Height + Gap + 60
main_frm.expirationdate_lbl.Left = 120 * 2
main_frm.expirationdate_lbl.Top = main_frm.stockonhold_lbl.Top + main_frm.stockonhold_lbl.Height + Gap + 60

main_frm.stockid_txtbox.Left = 120 + EditComponentGap
main_frm.stockid_txtbox.Top = main_frm.stockid_lbl.Top
main_frm.stockname_txtbox.Left = 120 + EditComponentGap
main_frm.stockname_txtbox.Top = main_frm.stockname_lbl.Top
main_frm.stockname_txtbox.Width = (StockGridWidth + 120) - (240 + EditComponentGap)
main_frm.categ_txtbox.Left = 120 + EditComponentGap
main_frm.categ_txtbox.Top = main_frm.categ_lbl.Top
main_frm.categ_txtbox.Width = ((StockGridWidth + 120) - (240 + EditComponentGap)) * 0.5
main_frm.addcateg_btn.Left = main_frm.categ_txtbox.Left + main_frm.categ_txtbox.Width + Gap
main_frm.addcateg_btn.Top = main_frm.categ_txtbox.Top
main_frm.addcateg_btn.Width = (main_frm.stockname_txtbox.Width * 0.25) - Gap
main_frm.addcateg_btn.Height = main_frm.categ_txtbox.Height
main_frm.delcateg_btn.Left = main_frm.addcateg_btn.Left + main_frm.addcateg_btn.Width + Gap
main_frm.delcateg_btn.Top = main_frm.categ_txtbox.Top
main_frm.delcateg_btn.Width = (main_frm.stockname_txtbox.Width * 0.25) - Gap
main_frm.delcateg_btn.Height = main_frm.categ_txtbox.Height
main_frm.price_txtbox.Left = 120 + EditComponentGap
main_frm.price_txtbox.Top = main_frm.price_lbl.Top
main_frm.price_txtbox.Width = (StockGridWidth + 120) - (240 + EditComponentGap)
main_frm.stockonhold_txtbox.Left = 120 + EditComponentGap
main_frm.stockonhold_txtbox.Top = main_frm.stockonhold_lbl.Top
main_frm.stockonhold_txtbox.Width = (StockGridWidth + 120) - (240 + EditComponentGap)
main_frm.month_cmbox.Left = 120 + EditComponentGap
main_frm.month_cmbox.Top = main_frm.expirationdate_lbl.Top
main_frm.month_cmbox.Width = (main_frm.stockonhold_txtbox.Width * 0.4) - (Gap * 2)
main_frm.day_cmbox.Left = main_frm.month_cmbox.Left + main_frm.month_cmbox.Width + Gap
main_frm.day_cmbox.Top = main_frm.month_cmbox.Top
main_frm.day_cmbox.Width = (main_frm.stockonhold_txtbox.Width * 0.3)
main_frm.year_cmbox.Left = main_frm.day_cmbox.Left + main_frm.day_cmbox.Width + Gap
main_frm.year_cmbox.Top = main_frm.month_cmbox.Top
main_frm.year_cmbox.Width = (main_frm.stockonhold_txtbox.Width * 0.3)
main_frm.minstock_lbl.Left = 120 * 2
main_frm.minstock_lbl.Top = main_frm.expirationdate_lbl.Top + main_frm.expirationdate_lbl.Height + Gap + 60
main_frm.minstock_txtbox.Left = 120 + EditComponentGap
main_frm.minstock_txtbox.Top = main_frm.minstock_lbl.Top
main_frm.minstock_txtbox.Width = ((main_frm.stockonhold_txtbox.Width - main_frm.maxstock_lbl.Width) - (Gap * 2)) / 2
main_frm.maxstock_lbl.Left = main_frm.minstock_txtbox.Left + main_frm.minstock_txtbox.Width + Gap
main_frm.maxstock_lbl.Top = main_frm.minstock_lbl.Top + 60
main_frm.maxstock_txtbox.Left = main_frm.maxstock_lbl.Left + main_frm.maxstock_lbl.Width + Gap
main_frm.maxstock_txtbox.Top = main_frm.maxstock_lbl.Top - 60
main_frm.maxstock_txtbox.Width = main_frm.minstock_txtbox.Width

main_frm.edit_box.Top = main_frm.stockid_txtbox.Top - Gap
main_frm.edit_box.Left = 120
main_frm.edit_box.Height = Screen.Height - main_frm.stockid_txtbox.Top - StatusBarHeight
main_frm.edit_box.Width = StockGridWidth

'Right Grid
GridSpace = Screen.Width - StockGridWidth - 360
main_frm.expense_btn.Width = (GridSpace * 0.2) - (Gap * 2)
main_frm.expense_btn.Left = Screen.Width - main_frm.expense_btn.Width - (Gap * 2)
main_frm.expense_btn.Top = Screen.Height - StatusBarHeight - BigButtonsHeight - (Gap * 2)
main_frm.expense_btn.Height = BigButtonsHeight
main_frm.report_btn.Width = (GridSpace * 0.2) - Gap
main_frm.report_btn.Left = main_frm.expense_btn.Left - main_frm.report_btn.Width - Gap
main_frm.report_btn.Top = main_frm.expense_btn.Top
main_frm.report_btn.Height = BigButtonsHeight
main_frm.delrec_btn.Width = (GridSpace * 0.2) - Gap
main_frm.delrec_btn.Left = main_frm.report_btn.Left - main_frm.delrec_btn.Width - Gap
main_frm.delrec_btn.Top = main_frm.expense_btn.Top
main_frm.delrec_btn.Height = BigButtonsHeight
main_frm.addrec_btn.Width = (GridSpace * 0.4) - Gap
main_frm.addrec_btn.Left = main_frm.delrec_btn.Left - main_frm.addrec_btn.Width - Gap
main_frm.addrec_btn.Top = main_frm.expense_btn.Top
main_frm.addrec_btn.Height = BigButtonsHeight
main_frm.log_box.Top = main_frm.addrec_btn.Top - Gap
main_frm.log_box.Left = StockGridWidth + 240
main_frm.log_box.Height = BigButtonsHeight + (Gap * 2)
main_frm.log_box.Width = GridSpace

'Day logs
main_frm.grantotal_shape.Width = (GridSpace * 0.2)
main_frm.grantotal_shape.Left = Screen.Width - main_frm.grantotal_shape.Width - Gap
main_frm.grantotal_shape.Top = main_frm.log_box.Top - SmallButtonsHeight - Gap
main_frm.grantotal_shape.Height = SmallButtonsHeight
main_frm.grandtotalcap_lbl.Left = main_frm.grantotal_shape.Left + Gap
main_frm.grandtotalcap_lbl.Top = main_frm.grantotal_shape.Top + (main_frm.grantotal_shape.Height / 2) - (main_frm.grandtotalcap_lbl.Height / 2)
main_frm.grandtotal_lbl.Left = main_frm.grandtotalcap_lbl.Left + main_frm.grandtotalcap_lbl.Width + Gap
main_frm.grandtotal_lbl.Top = main_frm.grandtotalcap_lbl.Top
main_frm.alllog_btn.Width = (GridSpace * 0.2) - Gap
main_frm.alllog_btn.Left = main_frm.grantotal_shape.Left - main_frm.alllog_btn.Width - Gap
main_frm.alllog_btn.Top = main_frm.grantotal_shape.Top
main_frm.alllog_btn.Height = SmallButtonsHeight
main_frm.nextday_btn.Width = (GridSpace * 0.2) - Gap
main_frm.nextday_btn.Left = main_frm.alllog_btn.Left - main_frm.nextday_btn.Width - Gap
main_frm.nextday_btn.Top = main_frm.grantotal_shape.Top
main_frm.nextday_btn.Height = SmallButtonsHeight
main_frm.todaylog_btn.Width = (GridSpace * 0.2) - Gap
main_frm.todaylog_btn.Left = main_frm.nextday_btn.Left - main_frm.todaylog_btn.Width - Gap
main_frm.todaylog_btn.Top = main_frm.grantotal_shape.Top
main_frm.todaylog_btn.Height = SmallButtonsHeight
main_frm.prevday_btn.Width = (GridSpace * 0.2) - Gap
main_frm.prevday_btn.Left = main_frm.todaylog_btn.Left - main_frm.prevday_btn.Width - Gap
main_frm.prevday_btn.Top = main_frm.grantotal_shape.Top
main_frm.prevday_btn.Height = SmallButtonsHeight

'log eat all space
main_frm.logs_datgrid.Left = main_frm.stocks_datgrid.Width + 120 + Gap
main_frm.logs_datgrid.Top = main_frm.stocks_datgrid.Top
main_frm.logs_datgrid.Height = Screen.Height - TitleBarHeight - main_frm.log_box.Height - SmallButtonsHeight - StatusBarHeight - (Gap * 4)
main_frm.logs_datgrid.Width = GridSpace
'Statusbar
main_frm.StatusBar1.Height = StatusBarHeight
main_frm.StatusBar1.Panels.Add (1)
main_frm.StatusBar1.Panels.Add (2)
main_frm.StatusBar1.Panels.Add (3)
main_frm.StatusBar1.Panels(1).Width = Screen.Width * 0.2
main_frm.StatusBar1.Panels(2).Width = Screen.Width * 0.3
main_frm.StatusBar1.Panels(3).Width = Screen.Width * 0.5

End Sub

Public Sub SetupTitleBarForm()
GridSpace = Screen.Width - StockGridWidth - 360
Dim CenteredDate As Integer
CenteredDate = (GridSpace / 2) - ((titlebar_frm.to_lbl.Width + Gap + titlebar_frm.todate_lbl.Width + (Gap * 3) + titlebar_frm.from_lbl.Width + Gap + titlebar_frm.fromdate_lbl.Width) / 2) + StockGridWidth + (Gap * 2)

titlebar_frm.Height = TitleBarHeight
titlebar_frm.Width = Screen.Width

titlebar_frm.Top = 0
titlebar_frm.Left = 0

lR = SetTopMostWindow(titlebar_frm.hwnd, True)

titlebar_frm.settings_btn.Left = Gap
titlebar_frm.settings_btn.Top = Gap
titlebar_frm.exit_btn.Left = Screen.Width - titlebar_frm.exit_btn.Width - Gap
titlebar_frm.exit_btn.Top = Gap
titlebar_frm.help_btn.Left = titlebar_frm.exit_btn.Left - titlebar_frm.help_btn.Width - Gap
titlebar_frm.help_btn.Top = Gap

titlebar_frm.stocklist_lbl.Top = Gap
titlebar_frm.stocklist_lbl.Left = (StockGridWidth / 2) - (titlebar_frm.stocklist_lbl.Width / 2) + Gap
titlebar_frm.saleslog_lbl.Top = Gap
titlebar_frm.saleslog_lbl.Left = (GridSpace / 2) - (titlebar_frm.saleslog_lbl.Width / 2) + StockGridWidth + (Gap * 2)
'Search Component Layout
titlebar_frm.search_lbl.Left = 120
titlebar_frm.search_lbl.Top = TitleBarHeight - titlebar_frm.search_lbl.Height - Gap

titlebar_frm.search_txtbox.Left = titlebar_frm.search_lbl.Left + titlebar_frm.search_lbl.Width + Gap
titlebar_frm.search_txtbox.Top = titlebar_frm.search_lbl.Top - 60
titlebar_frm.search_txtbox.Width = ((StockGridWidth - titlebar_frm.search_lbl.Width - 120 - titlebar_frm.searchby_lbl.Width) * 0.6) - Gap

titlebar_frm.searchby_lbl.Left = titlebar_frm.search_txtbox.Left + titlebar_frm.search_txtbox.Width + Gap
titlebar_frm.searchby_lbl.Top = titlebar_frm.search_lbl.Top
titlebar_frm.searchby_cmbox.Left = titlebar_frm.searchby_lbl.Left + titlebar_frm.searchby_lbl.Width + Gap
titlebar_frm.searchby_cmbox.Top = titlebar_frm.search_txtbox.Top
titlebar_frm.searchby_cmbox.Width = ((StockGridWidth - titlebar_frm.search_lbl.Width - 120 - titlebar_frm.searchby_lbl.Width) * 0.4) - Gap

titlebar_frm.groupby_lbl.Left = StockGridWidth + (Gap * 2)
titlebar_frm.groupby_lbl.Top = TitleBarHeight - titlebar_frm.groupby_lbl.Height - Gap
titlebar_frm.groupby_cmbox.Left = titlebar_frm.groupby_lbl.Left + titlebar_frm.groupby_lbl.Width + Gap
titlebar_frm.groupby_cmbox.Top = TitleBarHeight - titlebar_frm.groupby_lbl.Height - Gap - 60
titlebar_frm.datefilter_btn.Width = titlebar_frm.groupby_cmbox.Width
titlebar_frm.datefilter_btn.Left = Screen.Width - titlebar_frm.datefilter_btn.Width - Gap
titlebar_frm.datefilter_btn.Top = TitleBarHeight - (SmallButtonsHeight - 125) - Gap + 60
titlebar_frm.datefilter_btn.Height = SmallButtonsHeight - 125

titlebar_frm.from_lbl.Left = CenteredDate
titlebar_frm.from_lbl.Top = titlebar_frm.groupby_lbl.Top
titlebar_frm.fromdate_lbl.Left = titlebar_frm.from_lbl.Left + Gap + titlebar_frm.from_lbl.Width
titlebar_frm.fromdate_lbl.Top = titlebar_frm.from_lbl.Top
titlebar_frm.to_lbl.Left = titlebar_frm.fromdate_lbl.Left + (Gap * 3) + titlebar_frm.fromdate_lbl.Width
titlebar_frm.to_lbl.Top = titlebar_frm.from_lbl.Top
titlebar_frm.todate_lbl.Left = titlebar_frm.to_lbl.Left + Gap + titlebar_frm.to_lbl.Width
titlebar_frm.todate_lbl.Top = titlebar_frm.from_lbl.Top

'Others
Call SetButtons

counter = 0
While counter < stockrs.Fields.Count
    titlebar_frm.searchby_cmbox.AddItem stockrs.Fields(counter).Name
    counter = counter + 1
Wend

titlebar_frm.searchby_cmbox.Text = titlebar_frm.searchby_cmbox.List(1)
titlebar_frm.searchby_cmbox.RemoveItem 5
End Sub

Public Sub SetupAddingStockForm()
addingstock_frm.stockname_txtbox.TabIndex = 0
addingstock_frm.category_cmbox.TabIndex = 1
addingstock_frm.price_txtbox.TabIndex = 2
addingstock_frm.numberofstocks_txtbox.TabIndex = 3
addingstock_frm.add_btn.TabIndex = 4
End Sub

Public Sub SetupLoggingForm()
Call SetupShadow.SetupForm(logging_frm)

Call ClearOrders

Call ProcessOn
Call logging_frm.RefreshGrid

If GetOrderNumber <> "" Then
    logging_frm.orderno_lbl = GetOrderNumber
Else
    logging_frm.orderno_lbl = 1
End If
logging_frm.orderdate_lbl = titlebar_frm.todate_lbl
logging_frm.totalprice_lbl = "0.00"

'layout for logging form'
logging_frm.Height = main_frm.logs_datgrid.Height
logging_frm.Width = main_frm.logs_datgrid.Width
logging_frm.Left = main_frm.logs_datgrid.Left
logging_frm.Top = main_frm.logs_datgrid.Top

logging_frm.search_lbl.Left = 500
logging_frm.search_lbl.Top = 500
logging_frm.stockname_txtbox.Left = logging_frm.search_lbl.Left + logging_frm.search_lbl.Width + Gap
logging_frm.stockname_txtbox.Top = logging_frm.search_lbl.Top

logging_frm.orderno_lb.Top = logging_frm.search_lbl.Top
logging_frm.orderno_lb.Left = logging_frm.Width - (logging_frm.orderdate_lb.Width * 2) - 500
logging_frm.orderno_lbl.Top = logging_frm.search_lbl.Top
logging_frm.orderno_lbl.Left = logging_frm.orderno_lb.Left + logging_frm.orderno_lb.Width + Gap

logging_frm.orderdate_lb.Top = logging_frm.search_lbl.Top + logging_frm.orderno_lb.Height + Gap
logging_frm.orderdate_lb.Left = logging_frm.orderno_lb.Left
logging_frm.orderdate_lbl.Top = logging_frm.orderdate_lb.Top
logging_frm.orderdate_lbl.Left = logging_frm.orderdate_lb.Left + logging_frm.orderdate_lb.Width + Gap

logging_frm.stockname_lb.Left = logging_frm.search_lbl.Left
logging_frm.stockname_lb.Top = logging_frm.search_lbl.Top + logging_frm.search_lbl.Height + (Gap * 2)
logging_frm.category_lb.Left = logging_frm.search_lbl.Left
logging_frm.category_lb.Top = logging_frm.stockname_lb.Top + logging_frm.stockname_lb.Height + Gap
logging_frm.price_lb.Left = logging_frm.search_lbl.Left
logging_frm.price_lb.Top = logging_frm.category_lb.Top + logging_frm.category_lb.Height + Gap
logging_frm.numberofstocks_lb.Left = logging_frm.search_lbl.Left
logging_frm.numberofstocks_lb.Top = logging_frm.price_lb.Top + logging_frm.price_lb.Height + Gap

logging_frm.stockname_lbl.Left = logging_frm.search_lbl.Left + logging_frm.stockname_lb.Width + Gap
logging_frm.stockname_lbl.Top = logging_frm.stockname_lb.Top
logging_frm.stocktype_lbl.Left = logging_frm.search_lbl.Left + logging_frm.category_lb.Width + Gap
logging_frm.stocktype_lbl.Top = logging_frm.category_lb.Top
logging_frm.price_lbl.Left = logging_frm.search_lbl.Left + logging_frm.price_lb.Width + Gap
logging_frm.price_lbl.Top = logging_frm.price_lb.Top
logging_frm.numberofstocks_lbl.Left = logging_frm.search_lbl.Left + logging_frm.numberofstocks_lb.Width + Gap
logging_frm.numberofstocks_lbl.Top = logging_frm.numberofstocks_lb.Top

logging_frm.orders_datgrid.Left = logging_frm.search_lbl.Left
logging_frm.orders_datgrid.Top = logging_frm.numberofstocks_lb.Top + logging_frm.numberofstocks_lb.Height + Gap
logging_frm.orders_datgrid.Width = logging_frm.Width - 1000
logging_frm.orders_datgrid.Height = logging_frm.Height * 0.4

logging_frm.howmany_lb.Top = logging_frm.orders_datgrid.Top + logging_frm.orders_datgrid.Height + Gap
logging_frm.howmany_lb.Left = logging_frm.search_lbl.Left + (logging_frm.decrement_btn.Width / 2)

logging_frm.decrement_btn.Left = logging_frm.search_lbl.Left
logging_frm.decrement_btn.Top = logging_frm.howmany_lb.Top + logging_frm.howmany_lb.Height + Gap

logging_frm.increment_btn.Left = logging_frm.howmany_lb.Left + logging_frm.howmany_lb.Width - (logging_frm.increment_btn.Width / 2)
logging_frm.increment_btn.Top = logging_frm.decrement_btn.Top

logging_frm.howmany_lbl.Left = ((logging_frm.increment_btn.Left - (logging_frm.decrement_btn.Left + logging_frm.decrement_btn.Width)) / 2) - (logging_frm.howmany_lbl.Width / 2) + logging_frm.decrement_btn.Left + logging_frm.decrement_btn.Width
logging_frm.howmany_lbl.Top = logging_frm.decrement_btn.Top

logging_frm.addorder_btn.Left = logging_frm.search_lbl.Left
logging_frm.addorder_btn.Top = logging_frm.Height - logging_frm.addorder_btn.Height - Gap
logging_frm.delorder_btn.Left = logging_frm.addorder_btn.Left + logging_frm.addorder_btn.Width + Gap
logging_frm.delorder_btn.Top = logging_frm.addorder_btn.Top

logging_frm.cancel_btn.Left = logging_frm.Width - logging_frm.cancel_btn.Width - 500
logging_frm.cancel_btn.Top = logging_frm.addorder_btn.Top
logging_frm.process_btn.Left = logging_frm.cancel_btn.Left - logging_frm.process_btn.Width - Gap
logging_frm.process_btn.Top = logging_frm.addorder_btn.Top

logging_frm.totalprice_lb.Left = (logging_frm.Width / 2) - ((logging_frm.totalprice_lb.Width + Gap + logging_frm.totalprice_lbl.Width) / 2)
logging_frm.totalprice_lb.Top = logging_frm.Height - logging_frm.totalprice_lb.Height - Gap
logging_frm.totalprice_lbl.Left = logging_frm.totalprice_lb.Left + logging_frm.totalprice_lb.Width + Gap
logging_frm.totalprice_lbl.Top = logging_frm.Height - logging_frm.totalprice_lbl.Height - (Gap / 2)
End Sub

Public Sub SetupDataGrids(stocktable As Boolean, logstable As Boolean)
Dim ScrollWidth As Integer
Dim LogGridWidth As Integer
LogGridWidth = main_frm.logs_datgrid.Width
ScrollWidth = 300

If stocktable Then
    main_frm.stocks_datgrid.Columns(0).Width = 0
    main_frm.stocks_datgrid.Columns(1).Width = StockGridWidth * 0.7 - ScrollWidth
    main_frm.stocks_datgrid.Columns(2).Width = 0
    main_frm.stocks_datgrid.Columns(3).Width = StockGridWidth * 0.3 - ScrollWidth
    main_frm.stocks_datgrid.Columns(4).Width = 0
    main_frm.stocks_datgrid.Columns(5).Width = 0
    main_frm.stocks_datgrid.Columns(6).Width = 0
    main_frm.stocks_datgrid.Columns(7).Width = 0
    
    main_frm.stocks_datgrid.Columns(3).NumberFormat = "0.00"
    
    Call CenteredGridValues(main_frm.stocks_datgrid, 5)
End If

If logstable Then
    Dim columnName As String
    Dim ColumnWidth As Integer
    counter = 0
    
    While counter < main_frm.logs_datgrid.Columns.Count
        columnName = main_frm.logs_datgrid.Columns(counter).Caption
        
        If columnName = "StockName" Then
            ColumnWidth = LogGridWidth * 0.38
        ElseIf columnName = "DateOrdered" Then
            ColumnWidth = LogGridWidth * 0.15
        ElseIf columnName = "ItemSold" Then
            ColumnWidth = LogGridWidth * 0.38
        Else
            If main_frm.logs_datgrid.Columns.Count = 4 Then
                ColumnWidth = LogGridWidth * 0.25
            ElseIf main_frm.logs_datgrid.Columns.Count = 5 Then
                ColumnWidth = LogGridWidth * 0.13
            ElseIf main_frm.logs_datgrid.Columns.Count = 6 Then
                ColumnWidth = LogGridWidth * 0.1
            End If
        End If
        
        main_frm.logs_datgrid.Columns(counter).Width = ColumnWidth
        
        If (InStr(columnName, "Unit") = 1) Or (InStr(columnName, "Total") = 1) Then
            main_frm.logs_datgrid.Columns(counter).NumberFormat = "0.00"
        End If
        
        counter = counter + 1
    Wend
    
    Call CenteredGridValues(main_frm.logs_datgrid, main_frm.logs_datgrid.Columns.Count)
End If

Call expenses_frm.RefreshGrid
End Sub

Public Sub CenteredGridValues(datgrid As DataGrid, colcount As Integer)
counter = 0

While counter < colcount
    datgrid.Columns(counter).Alignment = dbgCenter
    counter = counter + 1
Wend
End Sub

Public Sub SetWarningForm(head As String, desc As String, okbtn As Boolean, okcap As String, cancelbtn As Boolean, cancelcap As String, Optional cap As String)
warning_frm.head_lbl = head
warning_frm.desc_lbl = desc
warning_frm.ok_btn.Visible = okbtn
warning_frm.ok_btn.Caption = okcap
warning_frm.cancel_btn.Visible = cancelbtn
warning_frm.cancel_btn.Caption = cancelcap
warning_frm.Caption = cap
warning_frm.Show
End Sub

Public Sub SetButtons()
Dim EditComponents, NotEditComponents As Boolean

If EditMode Or AddMode Then
    EditComponents = True
    NotEditComponents = False
Else
    EditComponents = False
    NotEditComponents = True
End If

main_frm.ok_btn.Enabled = EditComponents
main_frm.cancel_btn.Enabled = EditComponents
main_frm.stockname_txtbox.Enabled = EditComponents
main_frm.categ_txtbox.Enabled = EditComponents
main_frm.price_txtbox.Enabled = EditComponents
main_frm.stockonhold_txtbox.Enabled = EditComponents
main_frm.month_cmbox.Enabled = EditComponents
main_frm.day_cmbox.Enabled = EditComponents
main_frm.year_cmbox.Enabled = EditComponents
main_frm.addcateg_btn.Enabled = EditComponents
main_frm.delcateg_btn.Enabled = EditComponents
main_frm.minstock_txtbox.Enabled = EditComponents
main_frm.maxstock_txtbox.Enabled = EditComponents

main_frm.addstock_btn.Enabled = NotEditComponents
main_frm.editstock_btn.Enabled = NotEditComponents
main_frm.delstock_btn.Enabled = NotEditComponents
main_frm.outofstock_btn.Enabled = NotEditComponents
main_frm.refreshstock_btn.Enabled = NotEditComponents
main_frm.restock_btn.Enabled = NotEditComponents
main_frm.addrec_btn.Enabled = NotEditComponents
main_frm.delrec_btn.Enabled = NotEditComponents
main_frm.report_btn.Enabled = NotEditComponents
main_frm.prevday_btn.Enabled = NotEditComponents
main_frm.todaylog_btn.Enabled = NotEditComponents
main_frm.nextday_btn.Enabled = NotEditComponents
main_frm.alllog_btn.Enabled = NotEditComponents
main_frm.expense_btn.Enabled = NotEditComponents
titlebar_frm.groupby_cmbox.Enabled = NotEditComponents
titlebar_frm.settings_btn.Enabled = NotEditComponents
titlebar_frm.help_btn.Enabled = NotEditComponents
titlebar_frm.datefilter_btn.Enabled = NotEditComponents
titlebar_frm.search_txtbox.Enabled = NotEditComponents
End Sub

Public Sub SetStockInfo()
If stockrs.RecordCount <> 0 Then
    main_frm.stockid_txtbox = Format(stockrs.Fields(0).Value, "000")
    main_frm.stockname_txtbox.Text = stockrs.Fields(1).Value
    main_frm.categ_txtbox.Text = stockrs.Fields(2).Value
    main_frm.price_txtbox.Text = Format(stockrs.Fields(3).Value, "0.00")
    main_frm.stockonhold_txtbox.Text = stockrs.Fields(4).Value
    If stockrs.Fields(5).Value <> "" Then
        main_frm.month_cmbox.Text = MonthName(Month(DateValue(stockrs.Fields(5).Value)))
        If Len(Day(stockrs.Fields(5).Value)) = 1 Then
            main_frm.day_cmbox.Text = "0" & Day(stockrs.Fields(5).Value)
        Else
            main_frm.day_cmbox.Text = Day(stockrs.Fields(5).Value)
        End If
        main_frm.year_cmbox.Text = Year(stockrs.Fields(5).Value)
    Else
        main_frm.month_cmbox.Text = "N/A"
        main_frm.day_cmbox.Text = "N/A"
        main_frm.year_cmbox.Text = "N/A"
    End If
    main_frm.minstock_txtbox.Text = stockrs.Fields(6).Value
    main_frm.maxstock_txtbox.Text = stockrs.Fields(7).Value
    
Else
    main_frm.stockid_txtbox = "N/A"
    main_frm.stockname_txtbox.Text = "N/A"
    main_frm.categ_txtbox.Text = "N/A"
    main_frm.price_txtbox.Text = "N/A"
    main_frm.stockonhold_txtbox.Text = "N/A"
    main_frm.month_cmbox.Text = "N/A"
    main_frm.day_cmbox.Text = "N/A"
    main_frm.year_cmbox.Text = "N/A"
    main_frm.minstock_txtbox.Text = "N/A"
    main_frm.maxstock_txtbox.Text = "N/A"
End If
End Sub

Public Sub ProcessOn()
logging_frm.stockname_lbl = main_frm.stockname_txtbox.Text
logging_frm.stocktype_lbl = main_frm.categ_txtbox.Text
logging_frm.price_lbl = main_frm.price_txtbox.Text
logging_frm.numberofstocks_lbl = main_frm.stockonhold_txtbox.Text
logging_frm.howmany_lbl = "0"

logging_frm.price_lbl = Format(logging_frm.price_lbl, "0.00")
Call ButtonsBehavior(logging_frm)
End Sub

Public Sub ProcessOff()
logging_frm.stockname_lbl = ""
logging_frm.stocktype_lbl = ""
logging_frm.price_lbl = ""
logging_frm.numberofstocks_lbl = ""
logging_frm.howmany_lbl = "0"

Call ButtonsBehavior(logging_frm)
End Sub

Public Sub ButtonsBehavior(formname As Form)
If formname.Name = "logging_frm" Then
    If temporaryOrdersTB.RecordCount = 0 Then
        logging_frm.process_btn.Enabled = False
    Else
        logging_frm.process_btn.Enabled = True
    End If
    
    If logging_frm.howmany_lbl = "0" Then
        logging_frm.addorder_btn.Enabled = False
    Else
        logging_frm.addorder_btn.Enabled = True
    End If
    
    If logging_frm.stockname_lbl = "" Then
        logging_frm.increment_btn.Enabled = False
        logging_frm.decrement_btn.Enabled = False
    Else
        logging_frm.increment_btn.Enabled = True
        logging_frm.decrement_btn.Enabled = True
    End If
    
    If temporaryOrdersTB.RecordCount <> 0 Then
        logging_frm.delorder_btn.Enabled = True
    Else
        logging_frm.delorder_btn.Enabled = False
    End If
ElseIf formname.Name = "main_frm" Then
    Dim HasStock, HasRecord As Boolean
    If stockrs.RecordCount = 0 Then
        HasStock = False
    Else
        HasStock = True
    End If
    If ungrouprs.RecordCount = 0 Then
        HasRecord = False
    Else
        HasRecord = True
    End If
    main_frm.delstock_btn.Enabled = HasStock
    main_frm.editstock_btn.Enabled = HasStock
    main_frm.addrec_btn.Enabled = HasStock
    
    main_frm.delrec_btn.Enabled = HasRecord
    
    If ExpensesTB.RecordCount <> 0 Then
        expenses_frm.delexp_btn.Enabled = True
    Else
        expenses_frm.delexp_btn.Enabled = False
    End If
End If
End Sub

Public Sub SetCategories(cmbox As ComboBox)
counter = 0

If CategoriesTB.RecordCount <> 0 Then
    CategoriesTB.MoveFirst
    While counter < CategoriesTB.RecordCount
        cmbox.AddItem CategoriesTB.Fields(1).Value
        CategoriesTB.MoveNext
        counter = counter + 1
    Wend
    cmbox.AddItem "N/A"
    cmbox.Text = cmbox.List(0)
End If
End Sub

Public Sub SetupSuppProducts()
Set restock_frm.suppstock_datgrid.DataSource = SuppProductsTB
SuppProductsTB.filter = "SupplierID LIKE '" & GetSuppID(restock_frm.suppname_txtbox.Text) & "'"
End Sub

Public Sub SetSuppliers()
counter = 0
restock_frm.supp_cmbox.Clear
If SuppliersTB.RecordCount <> 0 Then
    SuppliersTB.MoveFirst
    While counter < SuppliersTB.RecordCount
        restock_frm.supp_cmbox.AddItem SuppliersTB.Fields(1).Value
        SuppliersTB.MoveNext
        counter = counter + 1
    Wend
    If restock_frm.supp_cmbox.ListCount <> 0 Then
        restock_frm.supp_cmbox.Text = restock_frm.supp_cmbox.List(0)
    End If
End If
End Sub

Public Sub SetRestockProducts()
restockrs.Requery
If GetSuppID(restock_frm.supp_cmbox.Text) <> "" Then
    restockrs.filter = "SupplierID LIKE " & GetSuppID(restock_frm.supp_cmbox.Text)
End If
restock_frm.restock_datgrid.Columns(0).Width = 0
restock_frm.restock_datgrid.Columns(1).Width = (restock_frm.restock_datgrid.Width * 0.35) - Gap
restock_frm.restock_datgrid.Columns(1).Alignment = dbgCenter
restock_frm.restock_datgrid.Columns(2).Width = (restock_frm.restock_datgrid.Width * 0.35) - Gap
restock_frm.restock_datgrid.Columns(2).Alignment = dbgCenter
restock_frm.restock_datgrid.Columns(3).Width = (restock_frm.restock_datgrid.Width * 0.15) - Gap
restock_frm.restock_datgrid.Columns(3).Alignment = dbgCenter
restock_frm.restock_datgrid.Columns(3).NumberFormat = "0.00"
restock_frm.restock_datgrid.Columns(4).Width = (restock_frm.restock_datgrid.Width * 0.15) - Gap
restock_frm.restock_datgrid.Columns(4).Alignment = dbgCenter

If restockrs.RecordCount = 0 Then
    restock_frm.decrement_btn.Enabled = False
    restock_frm.increment_btn.Enabled = False
    restock_frm.restock_btn.Enabled = False
Else
    restock_frm.decrement_btn.Enabled = True
    restock_frm.increment_btn.Enabled = True
    restock_frm.restock_btn.Enabled = True
End If

restock_frm.howmany_lbl = "0"
End Sub
