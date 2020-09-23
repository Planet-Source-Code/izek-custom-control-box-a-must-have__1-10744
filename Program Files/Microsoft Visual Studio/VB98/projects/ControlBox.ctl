VERSION 5.00
Begin VB.UserControl controlbox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ScaleHeight     =   1995
   ScaleWidth      =   4680
   Begin Project1.drag drag1 
      Left            =   2160
      Top             =   720
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin VB.CommandButton MaxB 
      Caption         =   "R"
      Height          =   225
      Left            =   4200
      TabIndex        =   7
      Top             =   50
      Width           =   225
   End
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      Height          =   255
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton CloseB 
      Caption         =   "X"
      Height          =   225
      Left            =   4440
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Close"
      Top             =   50
      Width           =   225
   End
   Begin VB.CommandButton MinB 
      Caption         =   "M"
      Height          =   225
      Left            =   3960
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Minimize"
      Top             =   50
      Width           =   225
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      Picture         =   "ControlBox.ctx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   50
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Custom Control Box"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   300
      TabIndex        =   3
      Top             =   45
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Height          =   135
      Left            =   0
      TabIndex        =   5
      Top             =   160
      Width           =   3615
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu minimize 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu maximize 
         Caption         =   "M&aximize"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "controlbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub CloseB_Click()
End
End Sub
Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.SetFocus
End Sub
Private Sub CloseB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.SetFocus
End Sub
Private Sub exit_Click()
End
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'by adding this line of code u make form draggable
Call drag1.FormDrag(Form1)
End Sub

Private Sub MaxB_Click()
Call UserControl_Resize
If Form1.WindowState = 2 Then
    Form1.WindowState = 0
    Call UserControl_Resize
Else
    Form1.WindowState = 2
    Call UserControl_Resize
End If
End Sub
Private Sub MaxB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.SetFocus
End Sub

Private Sub maximize_Click()
Call UserControl_Resize
If Form1.WindowState = 2 Then
    Form1.WindowState = 0
    Call UserControl_Resize
Else
    Form1.WindowState = 2
    Call UserControl_Resize
End If
End Sub
Private Sub MinB_Click()
Form1.WindowState = 1
End Sub
Private Sub MinB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.SetFocus
End Sub
Private Sub minimize_Click()
Form1.WindowState = 1
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Call PopupMenu(menu)
End Sub
Private Sub UserControl_Resize()
On Error Resume Next
UserControl.Width = Form1.Width
On Error Resume Next
UserControl.Height = 300
On Error Resume Next
Label1.Width = Form1.Width
On Error Resume Next
Label2.Width = Form1.Width
On Error Resume Next
Label3.Width = Form1.Width
On Error Resume Next
CloseB.Left = Form1.Width - 240
On Error Resume Next
MaxB.Left = Form1.Width - 480
On Error Resume Next
MinB.Left = Form1.Width - 720
End Sub
Public Property Get Caption() As String
    Caption = Label1.Caption
End Property
Public Property Let Caption(ByVal NewCaption As String)
    Label1.Caption = NewCaption
End Property
Public Property Get Picture() As String
    Picture = Picture1.Picture
End Property
Public Property Let Picture(ByVal Picture As String)
    Picture1.Picture = LoadPicture(Picture)
End Property
Public Property Get MinCaption() As String
    MinCaption = MinB.Caption
End Property
Public Property Let MinCaption(ByVal NewCaption As String)
    MinB.Caption = NewCaption
End Property
Public Property Get MaxCaption() As String
    MaxCaption = MaxB.Caption
End Property
Public Property Let MaxCaption(ByVal NewCaption As String)
    MaxB.Caption = NewCaption
End Property
Public Property Get CloseCaption() As String
    CloseCaption = CloseB.Caption
End Property
Public Property Let CloseCaption(ByVal NewCaption As String)
    CloseB.Caption = NewCaption
End Property
Public Property Get FontBold() As Boolean
If Label1.FontBold = True Then
    FontBold = True
Else
    FontBold = False
End If
End Property
Public Property Get FontItalic() As Boolean
If Label1.FontItalic = True Then
    FontItalic = True
Else
    FontItalic = False
End If
End Property
Public Property Get FontName() As String
FontName = Label1.FontName
End Property
Public Property Get FontSize() As Integer
FontSize = Label1.FontSize
End Property
Public Property Get FontStrikethru() As Boolean
If Label1.FontStrikethru = True Then
    FontStrikethru = True
Else
    FontStrikethru = False
End If

End Property
Public Property Get FontUnderline() As Boolean
If Label1.FontUnderline = True Then
    FontUnderline = True
Else
    FontUnderline = False
End If
End Property
Public Property Let FontBold(bold As Boolean)
MinB.FontBold = bold
MaxB.FontBold = bold
CloseB.FontBold = bold
Label1.FontBold = bold
End Property
Public Property Let FontItalic(Italic As Boolean)
MinB.FontItalic = Italic
MaxB.FontItalic = Italic
CloseB.FontItalic = Italic
Label1.FontItalic = Italic
End Property
Public Property Let FontName(name As String)
MinB.FontName = name
MaxB.FontName = name
CloseB.FontName = name
Label1.FontName = name
End Property
Public Property Let FontSize(Size As Integer)
MinB.FontSize = Size
MaxB.FontSize = Size
CloseB.FontSize = Size
Label1.FontSize = Size
End Property
Public Property Let FontStrikethru(Strikethru As Boolean)
MinB.FontStrikethru = Strikethru
MaxB.FontStrikethru = Strikethru
CloseB.FontStrikethru = Strikethru
Label1.FontStrikethru = Strikethru
End Property
Public Property Let FontUnderline(Underline As Boolean)
MinB.FontUnderline = Underline
MaxB.FontUnderline = Underline
CloseB.FontUnderline = Underline
Label1.FontUnderline = Underline
End Property
Public Property Get EnableMinButton() As Boolean
    EnableMin = MinB.Enabled
End Property
Public Property Let EnableMinButton(enable As Boolean)
    MinB.Enabled = enable
End Property
Public Property Get EnableMaxButton() As Boolean
    EnableMax = MaxB.Enabled
End Property
Public Property Let EnableMaxButton(enable As Boolean)
    MaxB.Enabled = enable
End Property
Public Property Get EnableCloseButton() As Boolean
    EnableClose = CloseB.Enabled
End Property
Public Property Let EnableCloseButton(enable As Boolean)
    CloseB.Enabled = enable
End Property

