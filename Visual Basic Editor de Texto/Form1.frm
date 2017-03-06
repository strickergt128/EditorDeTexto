VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Limpiar Negrita, Cursiva y Subrayado"
      Height          =   735
      Left            =   5280
      TabIndex        =   8
      Top             =   6120
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2280
      List            =   "Form1.frx":000D
      TabIndex        =   7
      Text            =   "Estilo de Fuente"
      Top             =   480
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0030
      Left            =   4200
      List            =   "Form1.frx":0043
      TabIndex        =   6
      Text            =   "Tamaña de Fuente"
      Top             =   480
      Width           =   1695
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "Form1.frx":005B
      Left            =   8280
      List            =   "Form1.frx":0068
      TabIndex        =   5
      Text            =   "Color de Fondo de Objeto"
      Top             =   480
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form1.frx":0084
      Left            =   6120
      List            =   "Form1.frx":0091
      TabIndex        =   4
      Text            =   "Color de Fuente"
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "S"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "K"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "N"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   11055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   1455
      Left            =   360
      Top             =   5760
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1.Text = "10" Then
Text1.FontSize = "10"
End If
If Combo1.Text = "20" Then
Text1.FontSize = "20"
End If
If Combo1.Text = "30" Then
Text1.FontSize = "30"
End If
If Combo1.Text = "40" Then
Text1.FontSize = "40"
End If
If Combo1.Text = "50" Then
Text1.FontSize = "50"
End If
End Sub

Private Sub Combo2_Click()
If Combo2.Text = "Arial Black" Then
Text1.Font = "Arial Black"
End If
If Combo2.Text = "Arial" Then
Text1.Font = "Arial"
End If
If Combo2.Text = "Agency FB" Then
Text1.Font = "Agency FB"
End If
End Sub

Private Sub Combo3_Click()
If Combo3.Text = "Rojo" Then
Text1.ForeColor = vbRed
End If
If Combo3.Text = "Amarillo" Then
Text1.ForeColor = vbYellow
End If
If Combo3.Text = "Azul" Then
Text1.ForeColor = vbBlue
End If
End Sub

Private Sub Combo4_Click()
If Combo4.Text = "Verde" Then
Shape1.BackColor = vbGreen
End If
If Combo4.Text = "Naranja" Then
Shape1.BackColor = &H80FF&
End If
If Combo4.Text = "Rosado" Then
Shape1.BackColor = &HFF80FF
End If
End Sub

Private Sub Command1_Click()
Text1.FontBold = True
End Sub

Private Sub Command2_Click()
Text1.FontItalic = True
End Sub

Private Sub Command3_Click()
Text1.FontUnderline = True
End Sub

Private Sub Command4_Click()
Text1.FontBold = False
Text1.FontItalic = False
Text1.FontUnderline = False
End Sub
