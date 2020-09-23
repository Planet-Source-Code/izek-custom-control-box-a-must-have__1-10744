VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin Project1.controlbox controlbox1 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _extentx        =   8255
      _extenty        =   529
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
controlbox1.Caption = "My custom control box"
controlbox1.FontBold = True
controlbox1.FontItalic = True
controlbox1.FontSize = 10
controlbox1.FontStrikethru = True
controlbox1.FontUnderline = True
controlbox1.FontName = "arial"
End Sub
