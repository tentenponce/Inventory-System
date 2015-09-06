VERSION 5.00
Begin VB.Form addcategory_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ok_btn 
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox categ_txtbox 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      ToolTipText     =   "Stock Name"
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category Name"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1500
   End
End
Attribute VB_Name = "addcategory_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_btn_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call SetupShadow.SetupForm(Me)
End Sub

Private Sub ok_btn_Click()
If Me.categ_txtbox.Text <> "" Then
    CategoriesTB.filter = "Category LIKE '" & Me.categ_txtbox.Text & "'"
    If CategoriesTB.RecordCount = 0 Then
        CategoriesTB.AddNew
        CategoriesTB.Fields(1) = Me.categ_txtbox.Text
        CategoriesTB.Update
        CategoriesTB.Close
        CategoriesTB.Open "SELECT * FROM Categories"
        
        CategoriesTB.filter = adFilterNone
        CategoriesTB.Requery
        main_frm.categ_txtbox.AddItem Me.categ_txtbox.Text
        Unload Me
    Else
        Call SetWarningForm("Already Exist", Me.categ_txtbox.Text & " already exists on the category list.", False, "", True, "Ok")
    End If
Else
    Call SetWarningForm("Missing Field", "You cannot leave blank.", False, "", True, "Ok")
End If
CategoriesTB.filter = adFilterNone
End Sub
