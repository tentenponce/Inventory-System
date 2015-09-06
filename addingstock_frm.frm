VERSION 5.00
Begin VB.Form addingstock_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Add Stock"
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   ControlBox      =   0   'False
   Icon            =   "addingstock_frm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton delcateg_btn 
      BackColor       =   &H00FFFF80&
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
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton addcateg_btn 
      BackColor       =   &H00FFFF80&
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
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   720
      Width           =   495
   End
   Begin VB.ComboBox category_cmbox 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton add_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Add Stock"
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox stockname_txtbox 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      ToolTipText     =   "Stock Name"
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox numberofstocks_txtbox 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "Stock Type"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox price_txtbox 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Stock Type"
      Top             =   1200
      Width           =   855
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock On Hold:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1425
   End
End
Attribute VB_Name = "addingstock_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Shadow As clsShadow

Private Sub add_btn_Click()
If stockname_txtbox.Text = "" Or Me.category_cmbox.Text = "" Or Me.price_txtbox.Text = "" Or numberofstocks_txtbox.Text = "" Then
    Call SetWarningForm("Missing Fields", "You cannot leave blank.", False, "", True, "Ok")
Else
    stockrs.filter = stockrs.Fields(1).Name & " LIKE '" & stockname_txtbox.Text & "'"
    If stockrs.RecordCount = 0 Then
        If IsNumeric(price_txtbox.Text) And IsNumeric(numberofstocks_txtbox.Text) Then
            StocksTB.AddNew
            StocksTB.Fields(1) = stockname_txtbox.Text
            StocksTB.Fields(2) = price_txtbox.Text
            StocksTB.Fields(3) = numberofstocks_txtbox.Text
            StocksTB.Update
            StocksTB.Close
            StocksTB.Open "SELECT * FROM Stocks"
            
            StockCategoryTB.AddNew
            StockCategoryTB.Fields(0) = GetItemID(stockname_txtbox.Text)
            StockCategoryTB.Fields(1) = GetCategoryID(Me.category_cmbox.Text)
            StockCategoryTB.Update
            
            stockrs.filter = adFilterNone
            stockrs.Requery
            Call SetupDataGrids
            Call SetupModule.SetStockInfo
            Unload Me
        Else
            Call SetWarningForm("Content Mismatched", "Price and Stock On Hold must be numbers only.", False, "", True, "Ok")
        End If
    Else
        Call SetWarningForm("Already Exist", stockname_txtbox.Text & " already exists on the inventory.", False, "", True, "Ok")
    End If
    stockrs.filter = adFilterNone
    stockrs.Requery
    
    Call SetupDataGrids
End If
End Sub

Private Sub addcateg_btn_Click()
addcategory_frm.Show
Unload Me
End Sub

Private Sub cancel_btn_Click()
Unload Me
End Sub

Private Sub delcateg_btn_Click()
temprs.Open "SELECT Category FROM Categories AS a, StockCategory AS b WHERE a.CategoryID = b.CategoryID"
temprs.filter = "Category LIKE '" & Me.category_cmbox.Text & "'"
If temprs.RecordCount = 0 Then
    Conn.Execute "DELETE FROM Categories WHERE Category = '" & Me.category_cmbox.Text & "'"
    CategoriesTB.Requery
Else
    Call SetWarningForm("Item Category Conflict", "One of your stocks has this category, delete or edit them first.", False, "", True, "Ok")
End If
temprs.Close
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Call cancel_btn_Click
End If
End Sub

Private Sub Form_Load()
Call SetupShadow.SetupForm(Me)
Call SetupAddingStockForm
End Sub

