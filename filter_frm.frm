VERSION 5.00
Begin VB.Form filter_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Filter Log"
   ClientHeight    =   2355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton filter_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Filter"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cancel_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Cancel"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox tomonth_cmbox 
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
      ItemData        =   "filter_frm.frx":0000
      Left            =   840
      List            =   "filter_frm.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox today_cmbox 
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
      ItemData        =   "filter_frm.frx":008E
      Left            =   2160
      List            =   "filter_frm.frx":00EF
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox toyear_cmbox 
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
      ItemData        =   "filter_frm.frx":016F
      Left            =   3480
      List            =   "filter_frm.frx":0179
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
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
      ItemData        =   "filter_frm.frx":0189
      Left            =   3480
      List            =   "filter_frm.frx":0193
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
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
      ItemData        =   "filter_frm.frx":01A3
      Left            =   2160
      List            =   "filter_frm.frx":0204
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   1215
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
      ItemData        =   "filter_frm.frx":0284
      Left            =   840
      List            =   "filter_frm.frx":02AC
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Left            =   3720
      TabIndex        =   10
      Top             =   120
      Width           =   495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Day"
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
      Left            =   2400
      TabIndex        =   9
      Top             =   120
      Width           =   495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   120
      Width           =   540
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
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
      TabIndex        =   7
      Top             =   1200
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
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
      TabIndex        =   0
      Top             =   600
      Width           =   615
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "filter_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim monthh, dayy, yearr As String
Dim date1, date2 As Date
Public filterr As String

Public Sub cancel_btn_Click()
Unload Me
End Sub

Public Sub filter_btn_Click()
monthh = GetMonthNumber(frommonth_cmbox.Text)
dayy = fromday_cmbox.Text
yearr = fromyear_cmbox.Text
date1 = dayy & "/" & monthh & "/" & yearr

monthh = GetMonthNumber(tomonth_cmbox.Text)
dayy = today_cmbox.Text
yearr = toyear_cmbox.Text
date2 = dayy & "/" & monthh & "/" & yearr

Call Filtering(Format$(date1, "d/m/yyyy"), Format$(date2, "d/m/yyyy"))

Call GrandTotalPrice
Call CalculateCashOnHold
Call CalculateInitialMoney
Call ButtonsBehavior(main_frm)
Call SetupDataGrids(False, True)
LogsMoveLast
Unload Me
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Call cancel_btn_Click
ElseIf KeyCode = 13 Then
    If filter_btn.Enabled Then
        Call filter_frm.filter_btn_Click
    End If
End If
End Sub

Public Sub Form_Load()
Call SetupShadow.SetupForm(Me)
date1 = Format(titlebar_frm.fromdate_lbl, "dd/mm/yyyy")
date2 = Format(titlebar_frm.todate_lbl, "dd/mm/yyyy")
frommonth_cmbox.Text = MonthName(Month(DateValue(date1)), False)
If Len(Day(date1)) = 1 Then
    fromday_cmbox.Text = "0" & Day(date1)
Else
    fromday_cmbox.Text = Day(date1)
End If
fromyear_cmbox.Text = Year(date1)
tomonth_cmbox.Text = MonthName(Month(DateValue(date2)), False)
If Len(Day(date2)) = 1 Then
    today_cmbox.Text = "0" & Day(date2)
Else
    today_cmbox.Text = Day(date2)
End If
toyear_cmbox.Text = Year(date2)
End Sub
