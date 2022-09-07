VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8100
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   465
      Left            =   2610
      TabIndex        =   6
      Top             =   4635
      Width           =   1230
   End
   Begin VB.CommandButton cmdprom 
      Caption         =   "Promedio"
      Height          =   465
      Left            =   2610
      TabIndex        =   5
      Top             =   3375
      Width           =   1230
   End
   Begin VB.CommandButton cmdsumar 
      Caption         =   "Sumar"
      Height          =   465
      Left            =   2610
      TabIndex        =   4
      Top             =   2700
      Width           =   1230
   End
   Begin VB.ListBox lstedad 
      Height          =   2400
      ItemData        =   "edades.frx":0000
      Left            =   495
      List            =   "edades.frx":0002
      TabIndex        =   3
      Top             =   2700
      Width           =   1230
   End
   Begin VB.CommandButton cmdagregar 
      Caption         =   "Agregar"
      Height          =   465
      Left            =   1485
      TabIndex        =   2
      Top             =   1710
      Width           =   1230
   End
   Begin VB.TextBox txtedad 
      Height          =   465
      Left            =   2430
      TabIndex        =   0
      Top             =   675
      Width           =   1230
   End
   Begin VB.Label lbledad 
      Caption         =   "Edad"
      Height          =   465
      Left            =   450
      TabIndex        =   1
      Top             =   675
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdagregar_Click()
lstedad.AddItem txtedad.Text
End Sub

Private Sub cmdprom_Click()

End Sub

Private Sub cmdsumar_Click()
Dim s As Integer
s = 0
Dim x As Byte
For x = 0 To lstedad.ListCount - 1
 lstedad.ListIndex = x
 s = s + Val(lstedad.Text)
MsgBox "Suma:" & s
Next x
End Sub
