VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1140
      Left            =   1575
      TabIndex        =   0
      Top             =   765
      Width           =   2760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n1, n2 As Integer

Private Sub Command1_Click()
n1 = InputBox("Ingrese un número", "Número 1")
n2 = InputBox("Ingrese otro número para completar la operación", "Número 2")
MsgBox "La suma es " & n1 + n2, , "Respuesta"
End Sub
