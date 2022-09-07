VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Cmb1 
      Height          =   315
      ItemData        =   "Fuentescoloresestilos.frx":0000
      Left            =   3465
      List            =   "Fuentescoloresestilos.frx":0013
      TabIndex        =   3
      Text            =   "Colores"
      Top             =   270
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estilos"
      Height          =   1905
      Left            =   405
      TabIndex        =   2
      Top             =   2160
      Width           =   2310
      Begin VB.CheckBox Chk1 
         Caption         =   "Negrita"
         Height          =   465
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Width           =   1230
      End
      Begin VB.CheckBox Chk2 
         Caption         =   "Cursiva"
         Height          =   465
         Left            =   90
         TabIndex        =   5
         Top             =   810
         Width           =   1230
      End
      Begin VB.CheckBox Chk3 
         Caption         =   "Subrayado"
         Height          =   465
         Left            =   135
         TabIndex        =   4
         Top             =   1260
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fuente"
      Height          =   1860
      Left            =   360
      TabIndex        =   1
      Top             =   180
      Width           =   2310
      Begin VB.OptionButton Option1 
         Caption         =   "Arial"
         Height          =   465
         Left            =   180
         TabIndex        =   9
         Top             =   315
         Width           =   1230
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Times New Roman"
         Height          =   465
         Left            =   135
         TabIndex        =   8
         Top             =   765
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Comic Sans"
         Height          =   465
         Left            =   135
         TabIndex        =   7
         Top             =   1215
         Width           =   1230
      End
   End
   Begin VB.TextBox txt1 
      Height          =   3165
      Left            =   3150
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   855
      Width           =   3570
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Chk1_Click()
If Chk1.Value = 1 Then
    txt1.FontBold = True
End If
If Chk1.Value = 0 Then
    txt1.FontBold = False
End If
End Sub
Private Sub Chk2_Click()
If Chk2.Value = 1 Then
    txt1.FontItalic = True
End If
If Chk2.Value = 0 Then
    txt1.FontItalic = False
End If
End Sub
Private Sub Chk3_Click()
If Chk3.Value = 1 Then
    txt1.FontUnderline = True
End If
If Chk3.Value = 0 Then
    txt1.FontUnderline = False
End If
End Sub
Private Sub Cmb1_Click()
If Cmb1.Text = "Azul" Then
    txt1.ForeColor = vbBlue
End If
If Cmb1.Text = "Rojo" Then
    txt1.ForeColor = vbRed
End If
If Cmb1.Text = "Negro" Then
    txt1.ForeColor = vbBlack
End If
If Cmb1.Text = "Amarillo" Then
    txt1.ForeColor = vbYellow
End If
If Cmb1.Text = "Verde" Then
    txt1.ForeColor = vbGreen
End If
End Sub
Private Sub Option1_Click()
If Option1.Value = True Then
    txt1.Font = "Arial"
End If
End Sub
Private Sub Option2_Click()
If Option2.Value = True Then
    txt1.Font = "Times New Roman"
End If
End Sub
Private Sub Option3_Click()
If Option3.Value = True Then
    txt1.Font = "Comic Sans MS"
End If
End Sub
