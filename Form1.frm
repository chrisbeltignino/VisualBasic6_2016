VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdFactorizar 
      Caption         =   "Factorizar "
      Height          =   465
      Left            =   1305
      TabIndex        =   1
      Top             =   1125
      Width           =   1725
   End
   Begin VB.TextBox TxtN 
      Height          =   420
      Left            =   1395
      TabIndex        =   0
      Top             =   270
      Width           =   1725
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function Factoriar(N As Integer) As Long
Dim i As Byte
Dim F As Long
F = 1
For i = 1 To N
F = F * i
Next i
Factoriar = F
End Function

Private Sub CmdFactorizar_Click()
MsgBox "Factorial de " & TxtN.Text & "=" & Factoriar(Val(TxtN.Text))
End Sub
