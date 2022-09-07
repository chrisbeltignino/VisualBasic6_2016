VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   7485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   LinkTopic       =   "Form5"
   Moveable        =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "menu.frx":0000
   ScaleHeight     =   7485
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Añadir Alumnos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FFFF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Volver"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Video Explicativo del Programa"
      Height          =   855
      Left            =   4200
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   4365
      Left            =   2880
      Picture         =   "menu.frx":58D3
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3885
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form5.Hide
Form6.Show
End Sub

Private Sub Command2_Click()
Form5.Hide
Form7.Show
End Sub

Private Sub Command3_Click()
MsgBox ("Esta Función no esta Disponible")
End Sub

Private Sub Command4_Click()
Form5.Hide
Form3.Show

End Sub

Private Sub Command5_Click()
End
End Sub
