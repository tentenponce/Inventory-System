VERSION 5.00
Begin VB.Form titlebar_frm 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox searchby_cmbox 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "titlebar_frm.frx":0000
      Left            =   4080
      List            =   "titlebar_frm.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ComboBox groupby_cmbox 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "titlebar_frm.frx":0004
      Left            =   9120
      List            =   "titlebar_frm.frx":0014
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton settings_btn 
      Height          =   615
      Left            =   120
      Picture         =   "titlebar_frm.frx":0034
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton exit_btn 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14520
      Picture         =   "titlebar_frm.frx":23FB
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton help_btn 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13680
      Picture         =   "titlebar_frm.frx":4909
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton datefilter_btn 
      Caption         =   "&Filter"
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
      Left            =   11040
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox search_txtbox 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label searchby_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3120
      TabIndex        =   14
      Top             =   1080
      Width           =   240
   End
   Begin VB.Label groupby_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GROUP BY:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8160
      TabIndex        =   13
      Top             =   960
      Width           =   870
   End
   Begin VB.Label saleslog_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Log"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9960
      TabIndex        =   11
      Top             =   120
      Width           =   2040
   End
   Begin VB.Label stocklist_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock List"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2640
      TabIndex        =   10
      Top             =   120
      Width           =   2100
   End
   Begin VB.Label todate_lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "1/1/2014"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14160
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label fromdate_lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "1/1/2014"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12720
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label to_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TO: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   13680
      TabIndex        =   4
      Top             =   1080
      Width           =   315
   End
   Begin VB.Label from_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FROM: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12120
      TabIndex        =   3
      Top             =   1080
      Width           =   585
   End
   Begin VB.Label search_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   555
   End
End
Attribute VB_Name = "titlebar_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Shadow As clsShadow

Public Sub SetupForm()
Set Shadow = New clsShadow
Call Shadow.Shadow(Me)
End Sub

Public Sub datefilter_btn_Click()
filter_frm.Show
End Sub

Private Sub exit_btn_Click()
Call SetWarningForm("Exit System", "Exit System? (All datas are saved.)", True, "Exit", True, "Cancel", "ExitSystem")
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call KeyCodeModule.ShortCutProblems(KeyCode, Shift, main_frm)
End Sub

Public Sub Form_Load()
Call Me.SetupForm
Call SetupModule.SetupTitleBarForm
End Sub

Public Sub groupby_cmbox_Click()
If groupby_cmbox.Text = "Order" Then
    'ungrouprs.Open "SELECT * FROM OrderGroup"
    Set main_frm.logs_datgrid.DataSource = orderrs
ElseIf groupby_cmbox.Text = "Date" Then
    'ungrouprs.Open "SELECT * FROM DateGroup"
    Set main_frm.logs_datgrid.DataSource = daters
ElseIf groupby_cmbox.Text = "Item" Then
    'ungrouprs.Open "SELECT * FROM StockGroup"
    Set main_frm.logs_datgrid.DataSource = itemrs
Else
    'ungrouprs.Open "SELECT * FROM LogsPerItem"
    Set main_frm.logs_datgrid.DataSource = ungrouprs
End If
Call SetupDataGrids(False, True)
Timers.switch_mode1.Enabled = True
End Sub

Public Sub help_btn_Click()
help_frm.Show
End Sub

Public Sub search_txtbox_Change()
If search_txtbox.Text = "" Then
    stockrs.filter = adFilterNone
Else
Dim ColName As String
ColName = titlebar_frm.searchby_cmbox.Text
    If ColName = "StockName" Or ColName = "Category" Then
        stockrs.filter = ColName & " LIKE '" & search_txtbox.Text & "*'"
    ElseIf (ColName = "StockID" Or ColName = "UnitPrice" Or ColName = "StockOnHold" Or ColName = "MinStock" Or ColName = "MaxStock") And IsNumeric(Me.search_txtbox.Text) Then
        stockrs.filter = ColName & " LIKE " & search_txtbox.Text
    End If
End If

Call SetupDataGrids(True, False)
Call SetupModule.SetStockInfo
Call CheckOutOfStock
Call ButtonsBehavior(main_frm)
End Sub

Private Sub settings_btn_Click()
Conn.Execute "UPDATE MinMaxStock SET MinStock=1"
End Sub
