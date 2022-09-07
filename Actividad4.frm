VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   870
      Left            =   1395
      TabIndex        =   0
      Top             =   855
      Width           =   2760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bas As Double
Dim are As Double
Dim alt As Double

Private Sub Command1_Click()
bas = InputBox("¿Cual es la base del triangulo?")
alt = InputBox("¿Cual es la altura del triangulo?")
are = bas * alt / 2
MsgBox "EL area del triangulo es: " & are
End Sub
