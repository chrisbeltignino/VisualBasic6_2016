VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt3 
      Height          =   825
      Left            =   1890
      TabIndex        =   8
      Top             =   3330
      Width           =   1185
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "Dividir"
      Height          =   510
      Left            =   4545
      TabIndex        =   7
      Top             =   2925
      Width           =   1230
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "Multiplicar"
      Height          =   510
      Left            =   4545
      TabIndex        =   6
      Top             =   2205
      Width           =   1230
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Sumar"
      Height          =   510
      Left            =   4545
      TabIndex        =   5
      Top             =   765
      Width           =   1230
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Restar"
      Height          =   510
      Left            =   4545
      TabIndex        =   4
      Top             =   1485
      Width           =   1230
   End
   Begin VB.TextBox txt2 
      Height          =   690
      Left            =   1800
      TabIndex        =   1
      Top             =   1935
      Width           =   2175
   End
   Begin VB.TextBox txt1 
      Height          =   690
      Left            =   1800
      TabIndex        =   0
      Top             =   810
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Resultado"
      Height          =   780
      Left            =   720
      TabIndex        =   9
      Top             =   3375
      Width           =   915
   End
   Begin VB.Label lbl2 
      Caption         =   "Numero 2"
      Height          =   690
      Left            =   540
      TabIndex        =   3
      Top             =   1935
      Width           =   870
   End
   Begin VB.Label lbl1 
      Caption         =   "Numero 1"
      Height          =   690
      Left            =   540
      TabIndex        =   2
      Top             =   810
      Width           =   870
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmd1_Click()
txt3.Text = Val(txt1.Text) + Val(txt2.Text)
If txt1.Text = "" Then
    Else
    MsgBox ("Tienes que ingresar un numero")
End If
If txt2.Text = "" Then
    Else
    MsgBox ("Tienes que ingresar un numero")
End If
If IsNumeric(txt1.Text) = True Then
    Else
    MsgBox ("Tiene que ser un numero")
End If
If IsNumeric(txt2.Text) = True Then
    Else
    MsgBox ("Tiene que ser un numero")
End If
End Sub

Private Sub cmd2_Click()
txt3.Text = Val(txt1.Text) - Val(txt2.Text)
If txt1.Text = "" Then
    Else
    MsgBox ("Tienes que ingresar un numero")
End If
If txt2.Text = "" Then
    Else
    MsgBox ("Tienes que ingresar un numero")
End If
If IsNumeric(txt1.Text) = True Then
    Else
    MsgBox ("Tiene que ser un numero")
End If
If IsNumeric(txt2.Text) = True Then
    Else
    MsgBox ("Tiene que ser un numero")
End If
End Sub

Private Sub cmd3_Click()
txt3.Text = Val(txt1.Text) * Val(txt2.Text)
If txt1.Text = "" Then
    Else
    MsgBox ("Tienes que ingresar un numero")
End If
If txt2.Text = "" Then
    Else
    MsgBox ("Tienes que ingresar un numero")
End If
If IsNumeric(txt1.Text) = True Then
    Else
    MsgBox ("Tiene que ser un numero")
End If
If IsNumeric(txt2.Text) = True Then
    Else
    MsgBox ("Tiene que ser un numero")
End If
End Sub

Private Sub cmd4_Click()
txt3.Text = Val(txt1.Text) / Val(txt2.Text)
If txt1.Text = "" Then
    Else
    MsgBox ("Tienes que ingresar un numero")
End If
If txt2.Text = "" Then
    Else
    MsgBox ("Tienes que ingresar un numero")
End If
If IsNumeric(txt1.Text) = True Then
    Else
    MsgBox ("Tiene que ser un numero")
End If

If IsNumeric(txt2.Text) = True Then
    Else
    MsgBox ("Tiene que ser un numero")
End If
End Sub
