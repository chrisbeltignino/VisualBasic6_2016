VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Load"
      Height          =   510
      Left            =   1125
      TabIndex        =   3
      Top             =   1935
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show"
      Height          =   510
      Left            =   1125
      TabIndex        =   2
      Top             =   1305
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pasar"
      Height          =   510
      Left            =   2790
      TabIndex        =   1
      Top             =   900
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1395
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1725
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Form1.txtCopia.Text = Text1.Text
End Sub

Private Sub Command2_Click()
Form1.Show
End Sub

Private Sub Command3_Click()
Load Form1
End Sub
