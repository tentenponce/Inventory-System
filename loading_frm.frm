VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form loading_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Welcome!"
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   Icon            =   "loading_frm.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "loading_frm.frx":1EA66
   ScaleHeight     =   6019.418
   ScaleMode       =   0  'User
   ScaleWidth      =   6858.975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer progress_timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7320
      Top             =   120
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label tips_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All of the records that you are recording are based on filter."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1800
      TabIndex        =   8
      Top             =   4320
      Width           =   4725
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Did you know?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory System"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   5820
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Let our coding do the talking."
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Error 404"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bravo 2.4"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6000
      TabIndex        =   3
      Top             =   2400
      Width           =   990
   End
   Begin VB.Label info_lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7890
      TabIndex        =   2
      Top             =   4080
      Width           =   45
   End
   Begin VB.Label percent_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   3960
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   5145
      Left            =   0
      Picture         =   "loading_frm.frx":20616
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   8040
   End
End
Attribute VB_Name = "loading_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim percent As Integer
Dim Shadow As clsShadow

Public Sub SetupForm()
Set Shadow = New clsShadow
Call Shadow.Shadow(Me)
End Sub

Private Sub Form_Load()
ProgressBar1.Min = 0
ProgressBar1.Max = 101
ProgressBar1.Value = 0
info_lbl = "Initializing System..."

Me.progress_timer.Enabled = True

Call Me.SetupForm
Call RandomTipsss
End Sub

Private Sub progress_timer_Timer()
percent_lbl = ProgressBar1.Value & "%"
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = 20 Then
    info_lbl = "Connecting to Database..."
    Call ConnectionModule.Main
ElseIf ProgressBar1.Value = 40 Then
    info_lbl = "Loading Records..."
    Load main_frm
ElseIf ProgressBar1.Value = 60 Then
    info_lbl = "Setting up System..."
    Load titlebar_frm
    
    Call Timers.AnimateMainForm
    
ElseIf ProgressBar1.Value = ProgressBar1.Max Then
    info_lbl = "Done!"
    
    Me.progress_timer.Enabled = False
    
    titlebar_frm.groupby_cmbox.Text = "Order"
    
    'System Default Setup
    Call SetupDataGrids(True, True)
    Call SubRoutinesModule.CheckOutOfStock
    Call SetButtons
    Call main_frm.todaylog_btn_Click
    Call GrandTotalPrice
    stockrs.Sort = "StockName ASC"
    ungrouprs.Sort = "OrderNo ASC"
    orderrs.Sort = "OrderNo ASC"
    Call SetStockInfo
    
    Call LogsMoveLast

    Timers.statusbar.Enabled = True
    
    Timers.loadingfrm_exit1.Enabled = True
    
    main_frm.addrec_btn.SetFocus
End If
End Sub
