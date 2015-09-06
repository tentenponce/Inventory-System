VERSION 5.00
Begin VB.Form SetupShadow 
   BorderStyle     =   0  'None
   Caption         =   "setupshadow"
   ClientHeight    =   195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   150
   LinkTopic       =   "Form1"
   ScaleHeight     =   195
   ScaleWidth      =   150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "SetupShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Shadow As clsShadow

Public Sub SetupForm(Form1 As Form)
Set Shadow = New clsShadow
Call Shadow.Shadow(Form1)
End Sub
