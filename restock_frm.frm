VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form restock_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton exit_btn 
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
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8040
      Width           =   1215
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
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton decrement_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8160
      Width           =   495
   End
   Begin VB.CommandButton increment_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8160
      Width           =   495
   End
   Begin MSDataGridLib.DataGrid restock_datgrid 
      Height          =   2775
      Left            =   240
      TabIndex        =   25
      Top             =   4800
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777088
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   29
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Height          =   855
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton savesupp_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton editsupp_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Edit Supplier"
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton addsupp_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Add Supplier"
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton delsupp_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Delete Supplier"
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox supp_cmbox 
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton delsuppprod_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Delete Supplying Product"
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
      TabIndex        =   15
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton addsuppprod_btn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Add Supplying Product"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox search_txtbox 
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
      Left            =   3120
      TabIndex        =   4
      Top             =   1680
      Width           =   3615
   End
   Begin MSDataGridLib.DataGrid suppstock_datgrid 
      Height          =   2535
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4471
      _Version        =   393216
      BackColor       =   16777088
      HeadLines       =   1
      RowHeight       =   29
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin VB.TextBox suppname_txtbox 
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
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      Height          =   3855
      Left            =   120
      Top             =   720
      Width           =   9135
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Supplied"
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
      Left            =   840
      TabIndex        =   29
      Top             =   7680
      Width           =   1725
   End
   Begin VB.Label howmany_lbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1575
      TabIndex        =   28
      Top             =   8160
      Width           =   225
   End
   Begin VB.Shape Shape2 
      Height          =   4095
      Left            =   120
      Top             =   4680
      Width           =   9135
   End
   Begin VB.Label suppid_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier ID:"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Left            =   7440
      TabIndex        =   21
      Top             =   1200
      Width           =   960
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier ID:"
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
      Left            =   7440
      TabIndex        =   20
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Suppliers:"
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
      TabIndex        =   17
      Top             =   360
      Width           =   825
   End
   Begin VB.Label Label7 
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
      Left            =   2400
      TabIndex        =   13
      Top             =   2400
      Width           =   990
   End
   Begin VB.Label Label5 
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
      Left            =   2400
      TabIndex        =   12
      Top             =   2760
      Width           =   765
   End
   Begin VB.Label stocktype_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   11
      Top             =   2760
      Width           =   45
   End
   Begin VB.Label Label4 
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
      Left            =   2400
      TabIndex        =   10
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label price_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13321
         SubFormatType   =   1
      EndProperty
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
      TabIndex        =   9
      Top             =   3120
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stocks Left:"
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
      TabIndex        =   8
      Top             =   3480
      Width           =   900
   End
   Begin VB.Label numberofstocks_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13321
         SubFormatType   =   1
      EndProperty
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
      Top             =   3480
      Width           =   45
   End
   Begin VB.Label stockname_lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   3600
      TabIndex        =   6
      Top             =   2400
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
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
      TabIndex        =   5
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Products Supplying:"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Name:"
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
      TabIndex        =   0
      Top             =   960
      Width           =   1245
   End
End
Attribute VB_Name = "restock_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adding, editing As Boolean

Public Sub addsupp_btn_Click()
Call EnableSupp
Me.suppname_txtbox.Text = ""
adding = True
editing = False
Me.suppid_lbl = "000"
End Sub

Public Sub DisableSupp()
Me.suppname_txtbox.Enabled = False
Me.search_txtbox.Enabled = False
Me.suppstock_datgrid.Enabled = False
Me.addsuppprod_btn.Enabled = False
Me.delsuppprod_btn.Enabled = False
Me.savesupp_btn.Enabled = False
Me.cancel_btn.Enabled = False

Me.supp_cmbox.Enabled = True
Me.addsupp_btn.Enabled = True
Me.editsupp_btn.Enabled = True
Me.delsupp_btn.Enabled = True
Me.stockname_lbl = "N/A"
Me.stocktype_lbl = "N/A"
Me.price_lbl = "N/A"
Me.numberofstocks_lbl = "N/A"
Me.suppid_lbl = "000"
End Sub

Public Sub EnableSupp()
Me.suppname_txtbox.Enabled = True
Me.search_txtbox.Enabled = True
Me.suppstock_datgrid.Enabled = True
Me.addsuppprod_btn.Enabled = True
Me.delsuppprod_btn.Enabled = True
Me.savesupp_btn.Enabled = True
Me.cancel_btn.Enabled = True

Me.supp_cmbox.Enabled = False
Me.addsupp_btn.Enabled = False
Me.editsupp_btn.Enabled = False
Me.delsupp_btn.Enabled = False
Call SetStockInfo
Call Me.SetStockInfos
End Sub

Public Sub addsuppprod_btn_Click()
Dim Exists As Boolean
If Me.stockname_lbl <> "N/A" Then
    counter = 0
    If TemporarySuppProductsTB.RecordCount <> 0 Then
        TemporarySuppProductsTB.MoveFirst
    End If
    While counter < TemporarySuppProductsTB.RecordCount
        If Me.stockname_lbl = TemporarySuppProductsTB.Fields(0).Value Then
            Call SetWarningForm("Already Exist", TemporarySuppProductsTB.Fields(0).Value & " already exists on the table.", False, "", True, "Ok")
            Exists = True
        End If
        counter = counter + 1
        TemporarySuppProductsTB.MoveNext
    Wend
    If Not Exists Then
        TemporarySuppProductsTB.AddNew
        TemporarySuppProductsTB.Fields(0) = Me.stockname_lbl
        TemporarySuppProductsTB.Update
        TemporarySuppProductsTB.Requery
        TemporarySuppProductsTB.MoveLast
    End If
Else
    Call SetWarningForm("Select Item", "Please select an item to be added as supplier's item.", False, "", True, "Ok")
End If
End Sub

Public Sub cancel_btn_Click()
Call DisableSupp
Call ClearSuppEdit
Call SetSuppliers
Call SetRestockProducts
End Sub

Public Sub decrement_btn_Click()
If howmany_lbl <> 0 Then
    howmany_lbl = howmany_lbl - 1
End If
End Sub

Public Sub delsupp_btn_Click()
Call SetWarningForm("Delete Supplier", "Are you sure you want to delete supplier " & restock_frm.supp_cmbox.Text & "?", True, "Delete", True, "Cancel", "DeleteSupp")
Call SetRestockProducts
End Sub

Public Sub delsuppprod_btn_Click()
If Not TemporarySuppProductsTB.EOF And Not TemporarySuppProductsTB.BOF Then
    Conn.Execute "DELETE * FROM TemporarySuppProducts WHERE StockName='" & TemporarySuppProductsTB.Fields(0).Value & "'"
    TemporarySuppProductsTB.Requery
Else
    Call SetWarningForm("No Selected Stock", "Select Stock on the table to be remove.", False, "", True, "Ok")
End If
End Sub

Public Sub editsupp_btn_Click()
If Me.supp_cmbox.Text <> "" Then
    Call ClearSuppEdit
    adding = False
    editing = True
    Call EnableSupp
    Me.suppname_txtbox.Enabled = False
    Me.suppname_txtbox.Text = Me.supp_cmbox.Text
    Me.suppid_lbl = Format(GetSuppID(Me.supp_cmbox.Text), "000")
    SuppProductsTB.filter = "SupplierID LIKE " & Me.suppid_lbl
    SuppProductsTB.MoveFirst
    counter = 0
    While counter < SuppProductsTB.RecordCount
        StocksTB.filter = "StockID LIKE " & SuppProductsTB.Fields(1).Value
        TemporarySuppProductsTB.AddNew
        TemporarySuppProductsTB.Fields(0) = StocksTB.Fields(1).Value
        TemporarySuppProductsTB.Update
        counter = counter + 1
        SuppProductsTB.MoveNext
    Wend
    TemporarySuppProductsTB.Requery
Else
    Call SetWarningForm("Select Supplier", "Please select supplier to be edit.", False, "", True, "Ok")
End If
End Sub

Private Sub exit_btn_Click()
Unload Me
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call ShortCutProblems(KeyCode, Shift, Me)
End Sub

Public Sub Form_Load()
Call SetupShadow.SetupForm(Me)
Set Me.suppstock_datgrid.DataSource = TemporarySuppProductsTB
Set Me.restock_datgrid.DataSource = restockrs
Call SetStockInfo
Call Me.SetStockInfos
Call SetSuppliers
Call ClearSuppEdit
Call SetRestockProducts
Call DisableSupp
End Sub

Public Sub increment_btn_Click()
howmany_lbl = howmany_lbl + 1
End Sub

Public Sub restock_btn_Click()
If howmany_lbl <> 0 Then
    stockrs.filter = "StockID LIKE " & GetItemID(restockrs.Fields(1).Value)
    If (Val(stockrs.Fields("StockOnHold").Value) + Val(restock_frm.howmany_lbl)) >= stockrs.Fields("MaxStock").Value Then
        stockrs.filter = adFilterNone
        Call SetWarningForm("OverStocking", "You reach maximum stock, you cannot order more.", False, "", True, "Ok")
    Else
        stockrs.filter = adFilterNone
        Call SetWarningForm("Restock", "You have ordered " & restock_frm.howmany_lbl & " pieces of " & restockrs.Fields(1).Value & ", are you sure?", True, "Yes", True, "Cancel", "Restock")
    End If
End If
End Sub

Public Sub savesupp_btn_Click()
Dim NameExists As Boolean
If adding Then
    If Me.suppname_txtbox.Text <> "" And TemporarySuppProductsTB.RecordCount <> 0 Then
        counter = 0
        While counter < Me.supp_cmbox.ListCount
            If Me.suppname_txtbox.Text = Me.supp_cmbox.List(counter) Then
                NameExists = True
            End If
            counter = counter + 1
        Wend
        
        If Not NameExists Then
            SuppliersTB.AddNew
            SuppliersTB.Fields(1) = Me.suppname_txtbox.Text
            SuppliersTB.Update
            
            counter = 0
            TemporarySuppProductsTB.MoveFirst
            
            While counter < TemporarySuppProductsTB.RecordCount
                SuppProductsTB.AddNew
                SuppProductsTB.Fields(0) = GetSuppID(Me.suppname_txtbox.Text)
                SuppProductsTB.Fields(1) = GetItemID(TemporarySuppProductsTB.Fields(0).Value)
                SuppProductsTB.Update
                counter = counter + 1
                TemporarySuppProductsTB.MoveNext
            Wend
            Call DisableSupp
            Call ClearSuppEdit
            Call SetSuppliers
            Call SetRestockProducts
        Else
            Call SetWarningForm("Already Exist", Me.suppname_txtbox.Text & " already exists on the suppliers list.", False, "", True, "Ok")
        End If
    Else
        Call SetWarningForm("Missing Fields", "Supplier Name and Products Supplied cannot be empty.", False, "", True, "Ok")
    End If
ElseIf editing Then
    If TemporarySuppProductsTB.RecordCount <> 0 Then
        Conn.Execute "DELETE * FROM SupplierProducts WHERE SupplierID=" & Me.suppid_lbl
        SuppProductsTB.Requery
        counter = 0
        TemporarySuppProductsTB.MoveFirst

        While counter < TemporarySuppProductsTB.RecordCount
            SuppProductsTB.AddNew
            SuppProductsTB.Fields(0) = Me.suppid_lbl
            SuppProductsTB.Fields(1) = GetItemID(TemporarySuppProductsTB.Fields(0).Value)
            SuppProductsTB.Update
            counter = counter + 1
            TemporarySuppProductsTB.MoveNext
        Wend
        Call DisableSupp
        Call ClearSuppEdit
        Call SetSuppliers
        SuppProductsTB.Requery
        Call SetRestockProducts
    Else
        Call SetWarningForm("Missing Fields", "Products Supplied cannot be empty.", False, "", True, "Ok")
    End If
End If
End Sub

Public Sub search_txtbox_Change()
If search_txtbox.Text = "" Then
    stockrs.filter = adFilterNone
Else
    stockrs.filter = "StockName LIKE '" & search_txtbox.Text & "*'"
End If

Call SetupDataGrids(True, False)
Call SetStockInfo
Call Me.SetStockInfos
Call CheckOutOfStock
Call ButtonsBehavior(main_frm)
End Sub

Public Sub SetStockInfos()
Me.stockname_lbl = main_frm.stockname_txtbox.Text
Me.stocktype_lbl = main_frm.categ_txtbox.Text
Me.price_lbl = main_frm.price_txtbox.Text
Me.numberofstocks_lbl = main_frm.stockonhold_txtbox.Text
End Sub

Public Sub supp_cmbox_Click()
Call SetRestockProducts
End Sub
