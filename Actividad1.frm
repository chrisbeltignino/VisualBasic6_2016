VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdate 
      Caption         =   "Atención"
      Height          =   510
      Left            =   2475
      TabIndex        =   3
      Top             =   1935
      Width           =   1230
   End
   Begin VB.CommandButton cmdex 
      Caption         =   "Exclamación"
      Height          =   510
      Left            =   765
      TabIndex        =   2
      Top             =   1935
      Width           =   1230
   End
   Begin VB.CommandButton cmdint 
      Caption         =   "Interrogación"
      Height          =   510
      Left            =   2475
      TabIndex        =   1
      Top             =   945
      Width           =   1230
   End
   Begin VB.CommandButton cmdinf 
      Caption         =   "Información"
      Height          =   510
      Left            =   765
      TabIndex        =   0
      Top             =   945
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdate_Click()
MsgBox "¡Algo anda mal!", vbCritical, "Atención"

End Sub

Private Sub cmdex_Click()
MsgBox "¡Vamo Bien!", vbExclamation, "Exclamación"
End Sub

Private Sub cmdinf_Click()
MsgBox "!Preste Atencion¡", vbInformation, "Información"
End Sub

Private Sub cmdint_Click()
MsgBox "¿Está Entendiendo?", vbQuestion, "Interrogación"
End Sub
