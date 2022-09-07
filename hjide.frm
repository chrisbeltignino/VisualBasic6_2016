VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCopia 
      Height          =   375
      Left            =   765
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1305
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Unload"
      Height          =   510
      Left            =   2790
      TabIndex        =   1
      Top             =   1665
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hide"
      Height          =   510
      Left            =   2790
      TabIndex        =   0
      Top             =   765
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Form1.Hide 'es igual a Me.Hide
End Sub

Private Sub Command2_Click()
Unload Form1 'unload me
End Sub
