VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmul 
      Caption         =   "Multiplicación"
      Height          =   465
      Left            =   2520
      TabIndex        =   3
      Top             =   2025
      Width           =   1230
   End
   Begin VB.CommandButton cmdsuma 
      Caption         =   "Suma"
      Height          =   465
      Left            =   675
      TabIndex        =   2
      Top             =   675
      Width           =   1230
   End
   Begin VB.CommandButton cmdresta 
      Caption         =   "Resta"
      Height          =   465
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   1230
   End
   Begin VB.CommandButton cmddiv 
      Caption         =   "División"
      Height          =   465
      Left            =   675
      TabIndex        =   0
      Top             =   1935
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n1, n2 As Single

Private Sub cmddiv_Click()
    n1 = InputBox("Ingrese un Número")
    n2 = InputBox("Ingrese otro Numero")
    MsgBox "La división = " & n1 / n2, vbInformation, "Resultado"
End Sub

Private Sub cmdmul_Click()
    n1 = InputBox("Ingrese un Número")
    n2 = InputBox("Ingrese otro Numero")
    MsgBox "La multiplicación = " & n1 * n2, vbInformation, "Resultado"
End Sub

Private Sub cmdresta_Click()
    n1 = InputBox("Ingrese un Número")
    n2 = InputBox("Ingrese otro Numero")
    MsgBox "La resta = " & n1 - n2, vbInformation, "Resultado"
End Sub

Private Sub cmdsuma_Click()
    n1 = InputBox("Ingrese un Número")
    n2 = InputBox("Ingrese otro Numero")
    MsgBox "La suma = " & n1 + n2, vbInformation, "Resultado"
End Sub
