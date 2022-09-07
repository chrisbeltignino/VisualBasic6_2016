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
   Begin VB.ComboBox CmbC 
      Height          =   315
      ItemData        =   "Comustibles.frx":0000
      Left            =   2835
      List            =   "Comustibles.frx":0010
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   630
      Width           =   1275
   End
   Begin VB.TextBox txtPLT 
      Height          =   285
      Left            =   4905
      TabIndex        =   7
      Top             =   675
      Width           =   1275
   End
   Begin VB.TextBox txtimpor 
      Height          =   510
      Left            =   2880
      TabIndex        =   2
      Top             =   2880
      Width           =   1230
   End
   Begin VB.TextBox txtlitro 
      Height          =   510
      Left            =   2835
      TabIndex        =   1
      Top             =   1980
      Width           =   1230
   End
   Begin VB.TextBox txttl 
      Height          =   510
      Left            =   2835
      TabIndex        =   0
      Text            =   "  "
      Top             =   1215
      Width           =   1230
   End
   Begin VB.Label Label4 
      Caption         =   "Importe $"
      Height          =   510
      Left            =   1260
      TabIndex        =   6
      Top             =   2880
      Width           =   1230
   End
   Begin VB.Label Label3 
      Caption         =   "$ Litro"
      Height          =   510
      Left            =   1260
      TabIndex        =   5
      Top             =   1980
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "Total/Litros"
      Height          =   510
      Left            =   1260
      TabIndex        =   4
      Top             =   1215
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "Combustible"
      Height          =   285
      Left            =   1305
      TabIndex        =   3
      Top             =   675
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtPLT_Change()
Select Case CmbC.Text
 Case "GasOil"
 txtPLT.Text = 17.5
Select Case CmbC.txt
 Case "GasOil Premium"
  txtPLT = 19.5
Select Case CmbC.Text
 Case "Nafta Super"
 txtPLT = 18
Select Case CmbC.Text
 Case "Nafta Premium"
 txtPLT = 20
End Sub

