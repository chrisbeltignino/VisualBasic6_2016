VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdC 
      Caption         =   "Cancelar"
      Height          =   465
      Left            =   3105
      TabIndex        =   8
      Top             =   2475
      Width           =   1365
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   7515
      TabIndex        =   6
      Top             =   810
      Width           =   1770
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   7560
      TabIndex        =   5
      Top             =   360
      Width           =   1770
   End
   Begin VB.TextBox TxtNombre 
      Height          =   330
      Left            =   1935
      TabIndex        =   3
      Top             =   585
      Width           =   825
   End
   Begin VB.TextBox TxtNumero 
      Height          =   330
      Left            =   990
      TabIndex        =   2
      Top             =   585
      Width           =   825
   End
   Begin VB.CommandButton CmdLlamar 
      Caption         =   "Llamar"
      Height          =   465
      Left            =   3150
      TabIndex        =   0
      Top             =   585
      Width           =   1185
   End
   Begin VB.Label LblHora 
      Caption         =   "Hora"
      Height          =   330
      Left            =   5940
      TabIndex        =   7
      Top             =   810
      Width           =   1410
   End
   Begin VB.Label LblFecha 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      Height          =   330
      Left            =   5940
      TabIndex        =   4
      Top             =   360
      Width           =   1410
   End
   Begin VB.Label LblClientes 
      Caption         =   "Clientes"
      Height          =   240
      Left            =   135
      TabIndex        =   1
      Top             =   630
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer


Private Sub CmdLlamar_Click()
For X = 0 To Form2.LstNumero.ListCount - 1
Form2.LstNumero.ListIndex = X
If TxtNumero.Text = Form2.LstNumero.Text Then
Form2.LstNombre.ListIndex = Form2.LstNumero.ListIndex
TxtNombre.Text = Form2.LstNombre.Text
End If
Next X
End Sub
