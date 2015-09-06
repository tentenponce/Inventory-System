VERSION 5.00
Begin VB.Form pos_frm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Point of Sales Form"
   ClientHeight    =   8790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cancel_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Exit"
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton print_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Print Report"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton refresh_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Refresh Report"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton tofiltered_btn 
      Height          =   255
      Left            =   3360
      Picture         =   "pos_frm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton tounfiltered_btn 
      Height          =   255
      Left            =   3360
      Picture         =   "pos_frm.frx":2CBA
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1800
      Width           =   495
   End
   Begin VB.ComboBox unfiltered_cmbox 
      Height          =   315
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox filtered_cmbox 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      Height          =   735
      Left            =   120
      Top             =   1680
      Width           =   7215
   End
   Begin VB.Label todate_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1/1/2015"
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
      Left            =   4560
      TabIndex        =   44
      Top             =   840
      Width           =   780
   End
   Begin VB.Label Label18 
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
      Left            =   4200
      TabIndex        =   43
      Top             =   840
      Width           =   240
   End
   Begin VB.Label fromdate_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1/1/2015"
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
      Left            =   2880
      TabIndex        =   42
      Top             =   840
      Width           =   780
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click arrows to add or remove category from filtered to unfiltered and vice versa."
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
      Left            =   480
      TabIndex        =   41
      Top             =   7920
      Width           =   6645
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIP:"
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
      Left            =   120
      TabIndex        =   40
      Top             =   7920
      Width           =   300
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Average Sales/Income (Per Day)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   39
      Top             =   6000
      Width           =   3945
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Sales/Income"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   38
      Top             =   4200
      Width           =   2385
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Statistics"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   37
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      Height          =   1335
      Left            =   120
      Top             =   6360
      Width           =   7215
   End
   Begin VB.Line Line4 
      X1              =   3480
      X2              =   3600
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line3 
      X1              =   3720
      X2              =   5160
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Label aveexp_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Average Exp"
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
      Left            =   3840
      TabIndex        =   36
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Average Expense:"
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
      Left            =   720
      TabIndex        =   35
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label aveinc_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Average Income:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      TabIndex        =   32
      Top             =   7200
      Width           =   2085
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Average Total Income:"
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
      Left            =   720
      TabIndex        =   31
      Top             =   7200
      Width           =   1830
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unfiltered Categories:"
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
      Left            =   3960
      TabIndex        =   30
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filtered Categories:"
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
      TabIndex        =   29
      Top             =   1920
      Width           =   1620
   End
   Begin VB.Line Line2 
      X1              =   3720
      X2              =   5160
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line1 
      X1              =   3480
      X2              =   3600
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Average Sales Income:"
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
      Left            =   720
      TabIndex        =   26
      Top             =   6480
      Width           =   1830
   End
   Begin VB.Label ave_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Average Income:"
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
      Left            =   3840
      TabIndex        =   25
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label totalinc_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Income"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      TabIndex        =   24
      Top             =   5400
      Width           =   1635
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Income:"
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
      Left            =   720
      TabIndex        =   23
      Top             =   5400
      Width           =   1110
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   120
      Top             =   4560
      Width           =   7215
   End
   Begin VB.Shape Shape1 
      Height          =   2415
      Left            =   120
      Top             =   1680
      Width           =   7215
   End
   Begin VB.Label topnumsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top Number of Stocks Sold:"
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
      Index           =   1
      Left            =   3600
      TabIndex        =   14
      Top             =   3000
      Width           =   2220
   End
   Begin VB.Label topnumsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top Number of Stocks Sold:"
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
      Index           =   0
      Left            =   3600
      TabIndex        =   4
      Top             =   2760
      Width           =   2220
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top Number of Stocks Sold:"
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
      Left            =   3600
      TabIndex        =   1
      Top             =   2520
      Width           =   2250
   End
   Begin VB.Label topnumsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "howmany"
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
      Index           =   9
      Left            =   5880
      TabIndex        =   22
      Top             =   3720
      Width           =   795
   End
   Begin VB.Label topnumsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "howmany"
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
      Index           =   8
      Left            =   5880
      TabIndex        =   21
      Top             =   3480
      Width           =   795
   End
   Begin VB.Label topnumsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "howmany"
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
      Index           =   7
      Left            =   5880
      TabIndex        =   20
      Top             =   3240
      Width           =   795
   End
   Begin VB.Label topnumsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "howmany"
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
      Index           =   6
      Left            =   5880
      TabIndex        =   19
      Top             =   3000
      Width           =   795
   End
   Begin VB.Label topnumsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "howmany"
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
      Index           =   5
      Left            =   5880
      TabIndex        =   18
      Top             =   2760
      Width           =   795
   End
   Begin VB.Label topnumsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top Number of Stocks Sold:"
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
      Index           =   4
      Left            =   3600
      TabIndex        =   17
      Top             =   3720
      Width           =   2220
   End
   Begin VB.Label topnumsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top Number of Stocks Sold:"
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
      Index           =   3
      Left            =   3600
      TabIndex        =   16
      Top             =   3480
      Width           =   2220
   End
   Begin VB.Label topnumsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top Number of Stocks Sold:"
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
      Index           =   2
      Left            =   3600
      TabIndex        =   15
      Top             =   3240
      Width           =   2220
   End
   Begin VB.Label topsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top Sold Item:"
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
      Index           =   4
      Left            =   720
      TabIndex        =   13
      Top             =   3720
      Width           =   1155
   End
   Begin VB.Label topsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top Sold Item:"
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
      Index           =   3
      Left            =   720
      TabIndex        =   12
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Label topsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top Sold Item:"
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
      Index           =   2
      Left            =   720
      TabIndex        =   11
      Top             =   3240
      Width           =   1155
   End
   Begin VB.Label topsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top Sold Item:"
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
      Index           =   1
      Left            =   720
      TabIndex        =   10
      Top             =   3000
      Width           =   1155
   End
   Begin VB.Label topsold_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top Sold Item:"
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
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   2760
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top 5 Sold Item:"
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
      Left            =   720
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label totalexp_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Expense"
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
      Left            =   3840
      TabIndex        =   9
      Top             =   4920
      Width           =   1125
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Expense:"
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
      Left            =   720
      TabIndex        =   8
      Top             =   4920
      Width           =   1185
   End
   Begin VB.Label salesincome_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Income:"
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
      Left            =   3840
      TabIndex        =   7
      Top             =   4680
      Width           =   1125
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Income:"
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
      Left            =   720
      TabIndex        =   6
      Top             =   4680
      Width           =   1110
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SALES REPORT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   3540
   End
   Begin VB.Label Label4 
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
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   465
   End
End
Attribute VB_Name = "pos_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cancel_btn_Click()
Unload Me
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call ShortCutProblems(KeyCode, Shift, Me)
End Sub

Public Sub Form_Load()
Call SetupShadow.SetupForm(Me)

fromdate_lbl = titlebar_frm.fromdate_lbl
todate_lbl = titlebar_frm.todate_lbl

Call GetTopSoldItem
Call GetTopNumSoldItem
Call SetCategories(pos_frm.filtered_cmbox)

Me.salesincome_lbl = main_frm.grandtotal_lbl
Me.totalexp_lbl = expenses_frm.exp_lbl
Me.totalinc_lbl = Format(Me.salesincome_lbl - Me.totalexp_lbl, "0.00")
Me.ave_lbl = Format(GetAverageIncome, "0.00")
Me.aveexp_lbl = Format(GetAverageExpense, "0.00")
Me.aveinc_lbl = Format(Me.ave_lbl - Me.aveexp_lbl, "0.00")

If Me.unfiltered_cmbox.Text = "" Then
    Me.tofiltered_btn.Enabled = False
End If

If Me.filtered_cmbox.Text = "" Then
    Me.tounfiltered_btn.Enabled = False
End If
End Sub

Private Sub print_btn_Click()
DataReport1.Top = TitleBarHeight
DataReport1.Height = Screen.Height - TitleBarHeight
DataReport1.Width = Screen.Width
DataReport1.Show
End Sub

Public Sub refresh_btn_Click()
Call Me.Form_Load
End Sub

Public Sub tofiltered_btn_Click()
Me.filtered_cmbox.AddItem Me.unfiltered_cmbox.Text
Me.filtered_cmbox.Text = Me.unfiltered_cmbox.Text

Me.unfiltered_cmbox.RemoveItem Me.unfiltered_cmbox.ListIndex
If Me.unfiltered_cmbox.ListCount <> 0 Then
    Me.unfiltered_cmbox.Text = Me.unfiltered_cmbox.List(0)
Else
    Me.tofiltered_btn.Enabled = False
End If
Me.tounfiltered_btn.Enabled = True

Call GetTopSoldItem
Call GetTopNumSoldItem
End Sub

Public Sub tounfiltered_btn_Click()
Me.unfiltered_cmbox.AddItem Me.filtered_cmbox.Text
Me.unfiltered_cmbox.Text = Me.filtered_cmbox.Text

Me.filtered_cmbox.RemoveItem Me.filtered_cmbox.ListIndex
If Me.filtered_cmbox.ListCount <> 0 Then
    Me.filtered_cmbox.Text = Me.filtered_cmbox.List(0)
Else
    Me.tounfiltered_btn.Enabled = False
End If
Me.tofiltered_btn.Enabled = True

Call GetTopSoldItem
Call GetTopNumSoldItem
End Sub
