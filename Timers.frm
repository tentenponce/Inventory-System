VERSION 5.00
Begin VB.Form Timers 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer randomtips 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   840
      Top             =   2640
   End
   Begin VB.Timer mainfrm_close 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   120
   End
   Begin VB.Timer titlebarfrm_close 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   120
   End
   Begin VB.Timer switch_mode2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   720
   End
   Begin VB.Timer switch_mode1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   720
   End
   Begin VB.Timer expense_close 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   720
   End
   Begin VB.Timer expense_open 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   720
   End
   Begin VB.Timer loadingfrm_exit2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   120
   End
   Begin VB.Timer loadingfrm_exit1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer mainfrm_open 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer titlebar_open 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer statusbar 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2640
   End
End
Attribute VB_Name = "Timers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lR As Long

Private Sub expense_close_Timer()
If expenses_frm.Left < Screen.Width Then
    expenses_frm.Left = expenses_frm.Left + 100
    main_frm.expense_btn.Enabled = False
    expenses_frm.cancel_btn.Enabled = False
Else
    expenses_frm.Left = Screen.Width
    main_frm.expense_btn.Enabled = True
    expenses_frm.cancel_btn.Enabled = True
    main_frm.SetFocus
    expense_close.Enabled = False
End If
End Sub

Private Sub expense_open_Timer()
If expenses_frm.Left > (Screen.Width - expenses_frm.Width) Then
    expenses_frm.Left = expenses_frm.Left - 100
    main_frm.expense_btn.Enabled = False
    expenses_frm.cancel_btn.Enabled = False
Else
    expenses_frm.Left = Screen.Width - expenses_frm.Width
    main_frm.expense_btn.Enabled = True
    expenses_frm.cancel_btn.Enabled = True
    expense_open.Enabled = False
End If
End Sub

Private Sub loadingfrm_exit1_Timer()
If loading_frm.Top > 1500 Then
    loading_frm.Top = loading_frm.Top - 50
Else
    loadingfrm_exit1.Enabled = False
    loadingfrm_exit2.Enabled = True
End If
End Sub

Private Sub loadingfrm_exit2_Timer()
If loading_frm.Top <> Screen.Height Then
    loading_frm.Top = loading_frm.Top + 100
Else
Unload loading_frm
End If
End Sub

Private Sub mainfrm_close_Timer()
If main_frm.Top > -Screen.Height Then
    main_frm.Top = main_frm.Top - 200
Else
    lR = SetTopMostWindow(titlebar_frm.hwnd, False)
    main_frm.Top = -Screen.Height
    mainfrm_close.Enabled = False
    titlebarfrm_close.Enabled = True
    Me.expense_close.Enabled = True
End If
End Sub

Private Sub mainfrm_open_Timer()
If main_frm.Top < 0 Then
    main_frm.Top = main_frm.Top + 200
Else
    main_frm.Top = 0
    titlebar_frm.Enabled = True
    mainfrm_open.Enabled = False
    titlebar_frm.search_txtbox.SetFocus
End If
End Sub

Private Sub randomtips_Timer()
Randomize
main_frm.StatusBar1.Panels(3).Text = Randomtipss(Int(Rnd() * 19))
End Sub

Private Sub statusbar_Timer()
main_frm.StatusBar1.Panels(2).Text = DateTime.Now
main_frm.StatusBar1.Panels(1).Text = stockrs.AbsolutePosition & " of " & stockrs.RecordCount & " Records"
End Sub

Public Sub AnimateMainForm()
main_frm.Left = 0
main_frm.Top = -Screen.Height
main_frm.Width = Screen.Width
main_frm.Height = Screen.Height
titlebar_frm.Left = Screen.Width
main_frm.Show
titlebar_frm.Show
lR = SetTopMostWindow(titlebar_frm.hwnd, False)
titlebar_frm.Enabled = False
titlebar_open.Enabled = True
End Sub

Private Sub switch_mode1_Timer()
titlebar_frm.groupby_cmbox.Enabled = False
If main_frm.logs_datgrid.Width > 400 Then
    main_frm.logs_datgrid.Width = main_frm.logs_datgrid.Width - 400
    main_frm.logs_datgrid.Left = main_frm.logs_datgrid.Left + 200
Else
    main_frm.logs_datgrid.Width = 0
    Me.switch_mode2.Enabled = True
    Me.switch_mode1.Enabled = False
End If
End Sub

Private Sub switch_mode2_Timer()
If main_frm.logs_datgrid.Width < GridSpace Then
    main_frm.logs_datgrid.Width = main_frm.logs_datgrid.Width + 400
    main_frm.logs_datgrid.Left = main_frm.logs_datgrid.Left - 200
Else
    main_frm.logs_datgrid.Width = GridSpace
    Me.switch_mode2.Enabled = False
    main_frm.logs_datgrid.Left = main_frm.stocks_datgrid.Width + 120 + Gap
    titlebar_frm.groupby_cmbox.Enabled = True
End If
End Sub

Private Sub titlebar_open_Timer()
If titlebar_frm.Left > 0 Then
    titlebar_frm.Left = titlebar_frm.Left - 200
Else
    lR = SetTopMostWindow(titlebar_frm.hwnd, True)
    titlebar_frm.Left = 0
    titlebar_open.Enabled = False
    mainfrm_open.Enabled = True
End If
End Sub

Private Sub titlebarfrm_close_Timer()
If titlebar_frm.Left < Screen.Width Then
    titlebar_frm.Left = titlebar_frm.Left + 200
Else
    titlebarfrm_close.Enabled = False
    Unload main_frm
    Conn.Close
    Conn2.Close
    Unload DataReport1
    Unload addcategory_frm
    Unload addingexpense_frm
    Unload cashonhold_frm
    Unload expenses_frm
    Unload filter_frm
    Unload loading_frm
    Unload logging_frm
    Unload login_frm
    Unload main_frm
    Unload pos_frm
    Unload restock_frm
    Unload SetupShadow
    'Unload Timers
    Unload titlebar_frm
    Unload warning_frm
    Unload Me
End If
End Sub
