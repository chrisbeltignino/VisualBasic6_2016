VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7305
   LinkTopic       =   "Form2"
   ScaleHeight     =   5100
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstNombre 
      Height          =   1035
      ItemData        =   "lachonche.frx":0000
      Left            =   900
      List            =   "lachonche.frx":0013
      TabIndex        =   1
      Top             =   180
      Width           =   960
   End
   Begin VB.ListBox LstNumero 
      Height          =   1035
      ItemData        =   "lachonche.frx":003C
      Left            =   180
      List            =   "lachonche.frx":004F
      TabIndex        =   0
      Top             =   180
      Width           =   510
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

