VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9765
   ForeColor       =   &H0080FF80&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H008080FF&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2430
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   2070
      Width           =   1230
   End
   Begin VB.CommandButton cmdsaludo 
      Caption         =   "SALUDO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1125
      TabIndex        =   0
      Top             =   585
      Width           =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nombre As String

Private Sub cmdsalir_Click()
End
End Sub

Private Sub cmdsaludo_Click()
nombre = InputBox("ingrese Nombre", "Nombre")
MsgBox "Hola " & nombre, vbInformation, "Saludo"
End
End Sub
