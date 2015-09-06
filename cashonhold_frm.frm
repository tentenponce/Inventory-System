VERSION 5.00
Begin VB.Form cashonhold_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cashonhold_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Add as Cash On Hold"
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
      TabIndex        =   38
      Top             =   8445
      Width           =   1935
   End
   Begin VB.ComboBox frommonth_cmbox 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "cashonhold_frm.frx":0000
      Left            =   2040
      List            =   "cashonhold_frm.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   7200
      Width           =   1215
   End
   Begin VB.ComboBox fromday_cmbox 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "cashonhold_frm.frx":008E
      Left            =   3360
      List            =   "cashonhold_frm.frx":00EF
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   7200
      Width           =   1215
   End
   Begin VB.ComboBox fromyear_cmbox 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "cashonhold_frm.frx":016F
      Left            =   4680
      List            =   "cashonhold_frm.frx":0179
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton initial_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Add as Initial Money"
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8445
      Width           =   1935
   End
   Begin VB.CommandButton cancel_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Close (Esc)"
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8445
      Width           =   1815
   End
   Begin VB.TextBox a_txtbox 
      Height          =   375
      Left            =   1800
      TabIndex        =   17
      Top             =   5520
      Width           =   2535
   End
   Begin VB.TextBox e_txtbox 
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox ao_txtbox 
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox bo_txtbox 
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox eo_txtbox 
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox aoo_txtbox 
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox boo_txtbox 
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox eoo_txtbox 
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox aooo_txtbox 
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log Date:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   600
      TabIndex        =   37
      Top             =   7150
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   6240
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label total_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      TabIndex        =   31
      Top             =   6240
      Width           =   570
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3720
      TabIndex        =   30
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label a_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      TabIndex        =   29
      Top             =   5520
      Width           =   570
   End
   Begin VB.Label ao_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      TabIndex        =   28
      Top             =   4320
      Width           =   570
   End
   Begin VB.Label e_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      TabIndex        =   27
      Top             =   4920
      Width           =   570
   End
   Begin VB.Label bo_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      TabIndex        =   26
      Top             =   3720
      Width           =   570
   End
   Begin VB.Label eo_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      TabIndex        =   25
      Top             =   3120
      Width           =   570
   End
   Begin VB.Label boo_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      TabIndex        =   24
      Top             =   1920
      Width           =   570
   End
   Begin VB.Label aoo_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      TabIndex        =   23
      Top             =   2520
      Width           =   570
   End
   Begin VB.Label eoo_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      TabIndex        =   22
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label aooo_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      TabIndex        =   21
      Top             =   720
      Width           =   570
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pieces:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2520
      TabIndex        =   19
      Top             =   120
      Width           =   885
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Money:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   18
      Top             =   120
      Width           =   990
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   8
      Top             =   4920
      Width           =   570
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   7
      Top             =   5520
      Width           =   570
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   6
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   4
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "200.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1000.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "500.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   900
   End
End
Attribute VB_Name = "cashonhold_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public filter As String

Public Sub a_txtbox_Change()
Call MathProblems(Me.a_txtbox, Me.a_lbl, 1)
End Sub

Public Sub ao_txtbox_Change()
Call MathProblems(Me.ao_txtbox, Me.ao_lbl, 10)
End Sub

Public Sub aoo_txtbox_Change()
Call MathProblems(Me.aoo_txtbox, Me.aoo_lbl, 100)
End Sub

Public Sub aooo_txtbox_Change()
Call MathProblems(Me.aooo_txtbox, Me.aooo_lbl, 1000)
End Sub

Public Sub bo_txtbox_Change()
Call MathProblems(Me.bo_txtbox, Me.bo_lbl, 20)
End Sub

Public Sub boo_txtbox_Change()
Call MathProblems(Me.boo_txtbox, Me.boo_lbl, 200)
End Sub

Public Sub cancel_btn_Click()
Unload Me
End Sub

Private Sub cashonhold_btn_Click()
Dim total As Long
counter = 0

filter = CashOnHoldTB.filter
CashOnHoldTB.filter = adFilterNone

If CashOnHoldTB.RecordCount <> 0 Then
    CashOnHoldTB.MoveFirst
    While counter < CashOnHoldTB.RecordCount
        If GetDate(Me.fromday_cmbox.Text, GetMonthNumber(Me.frommonth_cmbox.Text), Me.fromyear_cmbox.Text) = CashOnHoldTB.Fields(1).Value Then
            CashOnHoldTB.Delete
            CashOnHoldTB.Requery
            counter = CashOnHoldTB.RecordCount
        Else
            counter = counter + 1
            CashOnHoldTB.MoveNext
        End If
    Wend
End If

CashOnHoldTB.AddNew
CashOnHoldTB.Fields(0) = Format(Me.total_lbl, "0.00")
CashOnHoldTB.Fields(1) = GetDate(Me.fromday_cmbox.Text, GetMonthNumber(Me.frommonth_cmbox.Text), Me.fromyear_cmbox.Text)
CashOnHoldTB.Update
If filter <> "0" Then
    CashOnHoldTB.filter = filter
End If

Call CalculateCashOnHold
End Sub

Public Sub e_txtbox_Change()
Call MathProblems(Me.e_txtbox, Me.e_lbl, 5)
End Sub

Public Sub eo_txtbox_Change()
Call MathProblems(Me.eo_txtbox, Me.eo_lbl, 50)
End Sub

Public Sub eoo_txtbox_Change()
Call MathProblems(Me.eoo_txtbox, Me.eoo_lbl, 500)
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Public Sub Form_Load()
Call SetupShadow.SetupForm(Me)
frommonth_cmbox.Text = MonthName(Month(DateValue(titlebar_frm.fromdate_lbl)), False)
If Len(Day(titlebar_frm.fromdate_lbl)) = 1 Then
    fromday_cmbox.Text = "0" & Day(titlebar_frm.fromdate_lbl)
Else
    fromday_cmbox.Text = Day(titlebar_frm.fromdate_lbl)
End If
fromyear_cmbox.Text = Year(titlebar_frm.fromdate_lbl)
End Sub

Private Sub initial_btn_Click()
Dim counter As Integer
Dim total As Long
counter = 0

filter = InitialMoneyTB.filter
InitialMoneyTB.filter = adFilterNone

If InitialMoneyTB.RecordCount <> 0 Then
    InitialMoneyTB.MoveFirst
    While counter < InitialMoneyTB.RecordCount
        If GetDate(Me.fromday_cmbox.Text, GetMonthNumber(Me.frommonth_cmbox.Text), Me.fromyear_cmbox.Text) = InitialMoneyTB.Fields(1).Value Then
            InitialMoneyTB.Delete
            InitialMoneyTB.Requery
            counter = InitialMoneyTB.RecordCount
        Else
            counter = counter + 1
            InitialMoneyTB.MoveNext
        End If
    Wend
End If

InitialMoneyTB.AddNew
InitialMoneyTB.Fields(0) = Format(Me.total_lbl, "0.00")
InitialMoneyTB.Fields(1) = GetDate(Me.fromday_cmbox.Text, GetMonthNumber(Me.frommonth_cmbox.Text), Me.fromyear_cmbox.Text)
InitialMoneyTB.Update
If filter <> "0" Then
    InitialMoneyTB.filter = filter
End If

Call CalculateInitialMoney
End Sub
