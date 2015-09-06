VERSION 5.00
Begin VB.Form addingexpense_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cancel_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox expprice_txtbox 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "Stock Type"
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox expdesc_txtbox 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   "Stock Type"
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton add_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Add Expense"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Description"
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
      Left            =   1080
      TabIndex        =   5
      Top             =   240
      Width           =   1980
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expensed Price"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   1320
      Width           =   1470
   End
End
Attribute VB_Name = "addingexpense_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub add_btn_Click()
If expprice_txtbox.Text <> "" And expdesc_txtbox.Text <> "" Then
    If IsNumeric(expprice_txtbox.Text) Then
        ExpensesTB.AddNew
        ExpensesTB.Fields(0) = DateValue(titlebar_frm.fromdate_lbl)
        ExpensesTB.Fields(1) = expdesc_txtbox.Text
        ExpensesTB.Fields(2) = Format(Val(expprice_txtbox.Text), "0.00")
        ExpensesTB.Update
        ExpensesTB.Requery
        'Call CalculateCashOnHold
        'Call CalculateInitialMoney
        Call ButtonsBehavior(main_frm)
        Call expenses_frm.RefreshGrid
        If ExpensesTB.RecordCount <> 0 Then
            ExpensesTB.MoveLast
        End If
        Unload Me
    Else
        Call SetWarningForm("Content Mismatched", "Price category must be numbers only.", False, "", True, "Ok")
    End If
Else
    Call SetWarningForm("Missing Fields", "You cannot leave blank.", False, "", True, "Ok")
End If
End Sub

Public Sub cancel_btn_Click()
Unload Me
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Call Me.cancel_btn_Click
End If
End Sub

Public Sub Form_Load()
Call SetupShadow.SetupForm(Me)
Me.expdesc_txtbox.TabIndex = 0
Me.expprice_txtbox.TabIndex = 1
Me.add_btn.TabIndex = 2
End Sub
