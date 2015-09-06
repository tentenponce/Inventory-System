VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form main_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   -375
   ClientTop       =   0
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton addrec_btn 
      Caption         =   "Add Record (F2)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      Picture         =   "main_frm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   9000
      Width           =   2055
   End
   Begin VB.ComboBox year_cmbox 
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
      ItemData        =   "main_frm.frx":2785
      Left            =   3840
      List            =   "main_frm.frx":27A1
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   9360
      Width           =   975
   End
   Begin VB.ComboBox day_cmbox 
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
      ItemData        =   "main_frm.frx":27D4
      Left            =   2760
      List            =   "main_frm.frx":2838
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   9360
      Width           =   975
   End
   Begin VB.ComboBox month_cmbox 
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
      ItemData        =   "main_frm.frx":28BD
      Left            =   1560
      List            =   "main_frm.frx":28E8
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   9360
      Width           =   1095
   End
   Begin VB.TextBox minstock_txtbox 
      Height          =   375
      Left            =   1560
      TabIndex        =   34
      Top             =   9840
      Width           =   495
   End
   Begin VB.TextBox maxstock_txtbox 
      Height          =   375
      Left            =   3720
      TabIndex        =   32
      Top             =   9840
      Width           =   495
   End
   Begin VB.CommandButton restock_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Restock"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Picture         =   "main_frm.frx":2953
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton expense_btn 
      Caption         =   "Expenses (F3)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11640
      Picture         =   "main_frm.frx":568E
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton refreshstock_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Refresh Stock"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "main_frm.frx":7D33
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton addcateg_btn 
      Caption         =   "Add"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8040
      Width           =   495
   End
   Begin VB.CommandButton delcateg_btn 
      Caption         =   "Remove"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8040
      Width           =   735
   End
   Begin VB.CommandButton outofstock_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Out of Stock"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      Picture         =   "main_frm.frx":9DC5
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton alllog_btn 
      Caption         =   "A&ll Logs"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton nextday_btn 
      Caption         =   "&Next Day Log"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      Picture         =   "main_frm.frx":CBC5
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton todaylog_btn 
      Caption         =   "&Today Log"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      Picture         =   "main_frm.frx":F878
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton prevday_btn 
      Caption         =   "&Previous Day Log"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Picture         =   "main_frm.frx":12537
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton report_btn 
      Caption         =   "Sales &Report"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   13440
      Picture         =   "main_frm.frx":151F1
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9360
      Width           =   1575
   End
   Begin VB.CommandButton delrec_btn 
      Caption         =   "&Delete Record"
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
      Left            =   7080
      Picture         =   "main_frm.frx":174B8
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   10080
      Width           =   855
   End
   Begin VB.CommandButton addstock_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Add Stock"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "main_frm.frx":19374
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5520
      Width           =   1215
   End
   Begin VB.ComboBox categ_txtbox 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton cancel_btn 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Picture         =   "main_frm.frx":1B44F
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10320
      Width           =   1335
   End
   Begin VB.CommandButton ok_btn 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Picture         =   "main_frm.frx":1EABA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   10320
      Width           =   1215
   End
   Begin VB.CommandButton editstock_btn 
      Caption         =   "&Edit Stock"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "main_frm.frx":21E69
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10320
      Width           =   1935
   End
   Begin VB.TextBox stockonhold_txtbox 
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   8880
      Width           =   3255
   End
   Begin VB.TextBox price_txtbox 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   8400
      Width           =   3255
   End
   Begin VB.TextBox stockname_txtbox 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   7560
      Width           =   3255
   End
   Begin VB.CommandButton delstock_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Delete Stock"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Picture         =   "main_frm.frx":24A4F
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   11025
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataGridLib.DataGrid stocks_datgrid 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777088
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   29
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid logs_datgrid 
      Height          =   6360
      Left            =   5280
      TabIndex        =   1
      Top             =   1080
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   11218
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777088
      HeadLines       =   1
      RowHeight       =   29
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label expirationdate_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ExpirationDate:"
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
      Left            =   120
      TabIndex        =   36
      Top             =   9360
      Width           =   1275
   End
   Begin VB.Label minstock_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Stock:"
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
      Left            =   120
      TabIndex        =   35
      Top             =   9960
      Width           =   1320
   End
   Begin VB.Label maxstock_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum Stock:"
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
      Left            =   2280
      TabIndex        =   33
      Top             =   9960
      Width           =   1350
   End
   Begin VB.Label grandtotal_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   13680
      TabIndex        =   26
      Top             =   7680
      Width           =   465
   End
   Begin VB.Shape grantotal_shape 
      Height          =   495
      Left            =   12120
      Top             =   7560
      Width           =   3015
   End
   Begin VB.Label grandtotalcap_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12240
      TabIndex        =   24
      Top             =   7680
      Width           =   600
   End
   Begin VB.Label stockid_txtbox 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock ID:"
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
      Left            =   1560
      TabIndex        =   12
      Top             =   7320
      Width           =   705
   End
   Begin VB.Label stockonhold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock On Hold:"
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
      Left            =   120
      TabIndex        =   11
      Top             =   8880
      Width           =   1200
   End
   Begin VB.Label price_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price:"
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
      Left            =   240
      TabIndex        =   9
      Top             =   8400
      Width           =   855
   End
   Begin VB.Label categ_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
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
      Left            =   240
      TabIndex        =   7
      Top             =   8040
      Width           =   765
   End
   Begin VB.Label stockid_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock ID:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   7320
      Width           =   705
   End
   Begin VB.Label stockname_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Name:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   7680
      Width           =   990
   End
   Begin VB.Shape edit_box 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      Height          =   3735
      Left            =   0
      Top             =   7200
      Width           =   4935
   End
   Begin VB.Shape log_box 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      Height          =   2655
      Left            =   5280
      Top             =   8160
      Width           =   9855
   End
End
Attribute VB_Name = "main_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub addcateg_btn_Click()
addcategory_frm.Show
End Sub

Public Sub addrec_btn_Click()
logging_frm.Show
End Sub

Public Sub addstock_btn_Click()
AddMode = True
Call SetupModule.SetButtons

main_frm.stockname_txtbox.Text = ""
main_frm.categ_txtbox.Text = main_frm.categ_txtbox.List(0)
main_frm.price_txtbox.Text = ""
main_frm.stockonhold_txtbox.Text = ""
main_frm.month_cmbox.Text = "N/A"
main_frm.day_cmbox.Text = "N/A"
main_frm.year_cmbox.Text = "N/A"
main_frm.minstock_txtbox.Text = "20"
main_frm.maxstock_txtbox.Text = "1000"
End Sub

Public Sub alllog_btn_Click()
daters.filter = adFilterNone
If daters.RecordCount <> 0 Then
    daters.MoveFirst
    Call Filtering(daters.Fields(0).Value, DateValue(Now))
Else
    Call Filtering("1/1/2014", DateValue(Now))
End If

ExpensesTB.Requery
Call GrandTotalPrice
Call CalculateCashOnHold
Call CalculateInitialMoney
Call LogsMoveLast
Call SetupDataGrids(False, True)
Call ButtonsBehavior(main_frm)
End Sub

Public Sub cancel_btn_Click()
EditMode = False
AddMode = False
Call SetupModule.SetButtons
Call ButtonsBehavior(main_frm)
Call SetStockInfo
End Sub

Public Sub delcateg_btn_Click()
temprs.Open "SELECT Category FROM Categories AS a, StockCategory AS b WHERE a.CategoryID = b.CategoryID"
temprs.filter = "Category LIKE '" & Me.categ_txtbox.Text & "'"
If temprs.RecordCount = 0 Then
    If main_frm.categ_txtbox.Text <> "N/A" Then
        Conn.Execute "DELETE * FROM Categories WHERE Category = '" & Me.categ_txtbox.Text & "'"
        CategoriesTB.Requery
        main_frm.categ_txtbox.RemoveItem categ_txtbox.ListIndex
    End If
Else
    Call SetWarningForm("Item Category Conflict", "One of your stocks has this category, delete or edit them first.", False, "", True, "Ok")
End If
temprs.Close
End Sub

Public Sub delrec_btn_Click()
If titlebar_frm.groupby_cmbox.Text = "Ungroup" Or titlebar_frm.groupby_cmbox.Text = "Order" Then
    Call SetWarningForm("Delete Order(s) Record", "Deleting this order " & ungrouprs.Fields(0).Value & " is not undoable, Are you sure you want to delete?", True, "Delete", True, "Cancel", "DeleteRec")
Else
    Call SetWarningForm("Read Only", "You can only delete on ungrouped table.", False, "", True, "Ok")
End If
End Sub

Public Sub delstock_btn_Click()
Call SetWarningForm("Delete Stock Record", "Deleting this stock will also cause deletion of it on the sales log, Are you sure you want to delete?", True, "Delete", True, "Cancel", "DeleteStock")
End Sub

Public Sub editstock_btn_Click()
EditMode = True
Call SetupModule.SetButtons

tempName = main_frm.stockname_txtbox.Text
tempCateg = main_frm.categ_txtbox.Text
tempPrice = main_frm.price_txtbox.Text
tempStockOnHold = main_frm.stockonhold_txtbox.Text

Me.ok_btn.SetFocus
End Sub

Public Sub expense_btn_Click()
ExpensesTB.filter = "ExpenseDate >= #" & titlebar_frm.fromdate_lbl & "# AND ExpenseDate <= #" & titlebar_frm.todate_lbl & "#"
InitialMoneyTB.filter = "LogDate >= #" & titlebar_frm.fromdate_lbl & "# AND LogDate <= #" & titlebar_frm.todate_lbl & "#"
CashOnHoldTB.filter = "LogDate >= #" & titlebar_frm.fromdate_lbl & "# AND LogDate <= #" & titlebar_frm.todate_lbl & "#"

Unload expenses_frm
expenses_frm.Show
Call CalculateCashOnHold
Call CalculateInitialMoney

If expenses_frm.Left = Screen.Width Then
    Timers.expense_open.Enabled = True
Else
    Timers.expense_close.Enabled = True
End If

expenses_frm.addexp_btn.SetFocus
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call KeyCodeModule.ShortCutProblems(KeyCode, Shift, Me)
End Sub

Public Sub Form_Load()
Call SetupModule.SetupMainForm
End Sub

Public Sub logs_datgrid_HeadClick(ByVal ColIndex As Integer)
Dim sortField As String
Dim sortString As String

sortField = Me.logs_datgrid.Columns(ColIndex).Caption

If titlebar_frm.groupby_cmbox.Text = "Ungroup" Then
    If InStr(ungrouprs.Sort, "Asc") Then
        sortString = sortField & " Desc"
    Else
        sortString = sortField & " Asc"
    End If
    ungrouprs.Sort = sortString
ElseIf titlebar_frm.groupby_cmbox.Text = "Order" Then
    If InStr(orderrs.Sort, "Asc") Then
        sortString = sortField & " Desc"
    Else
        sortString = sortField & " Asc"
    End If
    orderrs.Sort = sortString
ElseIf titlebar_frm.groupby_cmbox.Text = "Date" Then
    If InStr(daters.Sort, "Asc") Then
        sortString = sortField & " Desc"
    Else
        sortString = sortField & " Asc"
    End If
    daters.Sort = sortString
ElseIf titlebar_frm.groupby_cmbox.Text = "Item" Then
    If InStr(itemrs.Sort, "Asc") Then
        sortString = sortField & " Desc"
    Else
        sortString = sortField & " Asc"
    End If
    itemrs.Sort = sortString
End If
If (logs_datgrid.SelBookmarks.Count <> 0) Then
    logs_datgrid.SelBookmarks.Remove 0
End If
End Sub

Public Sub nextday_btn_Click()
Call Filtering(DateValue(titlebar_frm.todate_lbl) + 1, DateValue(titlebar_frm.todate_lbl) + 1)
ExpensesTB.Requery
Call GrandTotalPrice
Call CalculateCashOnHold
Call CalculateInitialMoney
Call LogsMoveLast
Call SetupDataGrids(False, True)
Call ButtonsBehavior(main_frm)
End Sub

Public Sub ok_btn_Click()
If EditMode Then
    EditMode = False
    Call SetupModule.SetButtons
    If stockname_txtbox.Text = "" Or Me.categ_txtbox.Text = "" Or Me.categ_txtbox.Text = "N/A" Or Me.price_txtbox.Text = "" Or stockonhold_txtbox.Text = "" Then
        Call SetWarningForm("Missing Fields", "You cannot leave blank.", False, "", True, "Ok")
    Else
        If IsNumeric(price_txtbox.Text) And IsNumeric(stockonhold_txtbox.Text) Then
            If Me.month_cmbox.Text = "N/A" Or Me.day_cmbox.Text = "N/A" Or Me.year_cmbox.Text = "N/A" Then
                Conn.Execute "UPDATE Stocks SET StockName = '" & main_frm.stockname_txtbox.Text & "', UnitPrice = " & main_frm.price_txtbox.Text & ", StockOnHold = " & main_frm.stockonhold_txtbox.Text & " WHERE StockID = " & main_frm.stockid_txtbox
                Conn.Execute "UPDATE StockExpiration SET ExpirationDate=NULL WHERE StockID=" & Me.stockid_txtbox
            Else
                Conn.Execute "UPDATE Stocks SET StockName = '" & main_frm.stockname_txtbox.Text & "', UnitPrice = " & main_frm.price_txtbox.Text & ", StockOnHold = " & main_frm.stockonhold_txtbox.Text & " WHERE StockID = " & main_frm.stockid_txtbox
                Conn.Execute "UPDATE StockExpiration SET ExpirationDate='" & GetDate(GetMonthNumber(main_frm.month_cmbox.Text), Me.day_cmbox.Text, Me.year_cmbox.Text) & "' WHERE StockID=" & Me.stockid_txtbox
            End If
            Conn.Execute "UPDATE StockCategory SET CategoryID =" & GetCategoryID(main_frm.categ_txtbox.Text) & " WHERE StockID = " & main_frm.stockid_txtbox
            Conn.Execute "UPDATE MinMaxStock SET MinStock =" & Me.minstock_txtbox.Text & ", MaxStock=" & Me.maxstock_txtbox.Text & " WHERE StockID=" & main_frm.stockid_txtbox
            ungrouprs.Requery
            orderrs.Requery
            itemrs.Requery
            daters.Requery
        Else
            Call SetWarningForm("Content Mismatched", "Price and Stock On Hold must be numbers only.", False, "", True, "Ok")
        End If
    End If
ElseIf AddMode Then
    AddMode = False
    Call SetupModule.SetButtons
    If stockname_txtbox.Text = "" Or Me.categ_txtbox.Text = "" Or Me.categ_txtbox.Text = "N/A" Or Me.price_txtbox.Text = "" Or stockonhold_txtbox.Text = "" Or Me.minstock_txtbox.Text = "" Or Me.maxstock_txtbox.Text = "" Then
        Call SetWarningForm("Missing Fields", "You cannot leave blank or not available(N/A).", False, "", True, "Ok")
    Else
        stockrs.filter = stockrs.Fields(1).Name & " LIKE '" & stockname_txtbox.Text & "'"
        If stockrs.RecordCount = 0 Then
            If IsNumeric(maxstock_txtbox.Text) And IsNumeric(minstock_txtbox.Text) And IsNumeric(price_txtbox.Text) And IsNumeric(stockonhold_txtbox.Text) Then
                StocksTB.AddNew
                StocksTB.Fields(1) = stockname_txtbox.Text
                StocksTB.Fields(2) = price_txtbox.Text
                StocksTB.Fields(3) = stockonhold_txtbox.Text
                StocksTB.Update
                
                Conn.Execute "INSERT INTO StockExpiration VALUES(" & GetItemID(stockname_txtbox.Text) & ", NULL)"
                
                If Me.month_cmbox.Text = "N/A" Or Me.day_cmbox.Text = "N/A" Or Me.year_cmbox.Text = "N/A" Then
                    Conn.Execute "UPDATE StockExpiration SET ExpirationDate=NULL WHERE StockID=" & GetItemID(Me.stockname_txtbox.Text)
                Else
                    Conn.Execute "UPDATE StockExpiration SET ExpirationDate='" & GetDate(GetMonthNumber(main_frm.month_cmbox.Text), Me.day_cmbox.Text, Me.year_cmbox.Text) & "' WHERE StockID=" & GetItemID(Me.stockname_txtbox.Text)
                End If
                
                MinMaxStockTB.AddNew
                MinMaxStockTB.Fields(0) = GetItemID(stockname_txtbox.Text)
                MinMaxStockTB.Fields(1) = Me.minstock_txtbox.Text
                MinMaxStockTB.Fields(2) = Me.maxstock_txtbox.Text
                MinMaxStockTB.Update
                
                StockCategoryTB.AddNew
                StockCategoryTB.Fields(0) = GetItemID(stockname_txtbox.Text)
                StockCategoryTB.Fields(1) = GetCategoryID(Me.categ_txtbox.Text)
                StockCategoryTB.Update
                
                
                
            Else
                Call SetWarningForm("Content Mismatched", "Price, Stock On Hold, Minimum and Maximum Stock must be numbers only.", False, "", True, "Ok")
            End If
        Else
            Call SetWarningForm("Already Exist", stockname_txtbox.Text & " already exists on the inventory.", False, "", True, "Ok")
        End If
    End If
End If

stockrs.filter = adFilterNone
stockrs.Requery
Call SetupModule.SetStockInfo
Call SetupModule.SetupDataGrids(True, True)
Call ButtonsBehavior(main_frm)
End Sub

Public Sub outofstock_btn_Click()
stockrs.filter = "StockOnHold LIKE 0"
Call ButtonsBehavior(main_frm)
End Sub

Public Sub prevday_btn_Click()
Call Filtering(DateValue(titlebar_frm.todate_lbl) - 1, DateValue(titlebar_frm.todate_lbl) - 1)
ExpensesTB.Requery
Call GrandTotalPrice
Call CalculateCashOnHold
Call CalculateInitialMoney
Call LogsMoveLast
Call SetupDataGrids(False, True)
Call ButtonsBehavior(main_frm)
End Sub

Public Sub refreshstock_btn_Click()
stockrs.filter = adFilterNone
stockrs.Requery
Call SetupModule.SetStockInfo
Call SetupModule.SetupDataGrids(True, False)
Call ButtonsBehavior(main_frm)
End Sub

Public Sub report_btn_Click()
pos_frm.Show
End Sub

Public Sub restock_btn_Click()
restock_frm.Show
If restock_frm.restock_btn.Enabled Then
    restock_frm.restock_btn.SetFocus
Else
End If
End Sub

Public Sub todaylog_btn_Click()
Call Filtering(DateValue(Now), DateValue(Now))
ExpensesTB.Requery
Call GrandTotalPrice
Call CalculateCashOnHold
Call CalculateInitialMoney
Call LogsMoveLast
Call SetupDataGrids(False, True)
Call ButtonsBehavior(main_frm)
End Sub
