VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ch6 
      Caption         =   "Estudiantes"
      Height          =   510
      Left            =   6345
      TabIndex        =   10
      Top             =   3825
      Width           =   1230
   End
   Begin VB.CheckBox ch5 
      Caption         =   "Jubilados"
      Height          =   510
      Left            =   6345
      TabIndex        =   9
      Top             =   3105
      Width           =   1230
   End
   Begin VB.ComboBox cb1 
      Height          =   315
      ItemData        =   "Agencia de Turismo.frx":0000
      Left            =   1575
      List            =   "Agencia de Turismo.frx":000D
      TabIndex        =   8
      Top             =   2070
      Width           =   1230
   End
   Begin VB.ComboBox cb2 
      Height          =   315
      ItemData        =   "Agencia de Turismo.frx":002C
      Left            =   4590
      List            =   "Agencia de Turismo.frx":0039
      TabIndex        =   7
      Top             =   2025
      Width           =   1230
   End
   Begin VB.CheckBox ch4 
      Caption         =   "Bodegas"
      Height          =   510
      Left            =   3015
      TabIndex        =   6
      Top             =   3735
      Width           =   1230
   End
   Begin VB.CheckBox ch3 
      Caption         =   "Safari"
      Height          =   510
      Left            =   1620
      TabIndex        =   5
      Top             =   3780
      Width           =   1230
   End
   Begin VB.CheckBox ch2 
      Caption         =   "City Tour"
      Height          =   510
      Left            =   3015
      TabIndex        =   4
      Top             =   3060
      Width           =   1230
   End
   Begin VB.CheckBox ch1 
      Caption         =   "4x4"
      Height          =   510
      Left            =   1575
      TabIndex        =   3
      Top             =   3060
      Width           =   1230
   End
   Begin VB.TextBox txt2 
      Height          =   510
      Left            =   4590
      TabIndex        =   2
      Top             =   1260
      Width           =   1230
   End
   Begin VB.TextBox txt1 
      Height          =   510
      Left            =   1530
      TabIndex        =   1
      Top             =   1260
      Width           =   1230
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Calcular"
      Height          =   510
      Left            =   4185
      TabIndex        =   0
      Top             =   5175
      Width           =   1230
   End
   Begin VB.Label Label13 
      Caption         =   "Agencia de Turismo"
      Height          =   465
      Left            =   3015
      TabIndex        =   23
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lbl3 
      Height          =   465
      Left            =   1620
      TabIndex        =   22
      Top             =   5985
      Width           =   1185
   End
   Begin VB.Label lbl2 
      Height          =   465
      Left            =   1620
      TabIndex        =   21
      Top             =   5265
      Width           =   1185
   End
   Begin VB.Label lbl1 
      Height          =   510
      Left            =   1620
      TabIndex        =   20
      Top             =   4545
      Width           =   1185
   End
   Begin VB.Label Label9 
      Caption         =   "Valor Total en Euros"
      Height          =   510
      Left            =   180
      TabIndex        =   19
      Top             =   5985
      Width           =   1230
   End
   Begin VB.Label Label8 
      Caption         =   "Valor Total en Dolares"
      Height          =   510
      Left            =   180
      TabIndex        =   18
      Top             =   5265
      Width           =   1230
   End
   Begin VB.Label Label7 
      Caption         =   "Valor Total en Pesos"
      Height          =   510
      Left            =   180
      TabIndex        =   17
      Top             =   4545
      Width           =   1230
   End
   Begin VB.Label Label6 
      Caption         =   "Descuentos"
      Height          =   510
      Left            =   4950
      TabIndex        =   16
      Top             =   3510
      Width           =   1230
   End
   Begin VB.Label Label5 
      Caption         =   "Excurciones"
      Height          =   510
      Left            =   180
      TabIndex        =   15
      Top             =   3465
      Width           =   1230
   End
   Begin VB.Label Label4 
      Caption         =   "Transporte"
      Height          =   510
      Left            =   3195
      TabIndex        =   14
      Top             =   1935
      Width           =   1230
   End
   Begin VB.Label Label3 
      Caption         =   "Destino"
      Height          =   510
      Left            =   135
      TabIndex        =   13
      Top             =   1980
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad de Pasajeros"
      Height          =   510
      Left            =   3195
      TabIndex        =   12
      Top             =   1260
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre:"
      Height          =   510
      Left            =   135
      TabIndex        =   11
      Top             =   1260
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As Integer
Dim B As Integer
Dim C As Integer
Dim D As Integer
Dim E As Integer
Dim F As Integer
Dim G As Integer
Dim H As Integer
Dim A1 As Integer
Dim B1 As Integer
Dim C1 As Integer
Dim D1 As Integer
Dim A2 As Integer
Dim B2 As Integer
Dim Subt As Integer
Dim Total As Integer
Dim Por As Integer

Private Sub ch1_Click()
A1 = 300
End Sub

Private Sub ch2_Click()
B1 = 200
End Sub

Private Sub ch3_Click()
C1 = 500
End Sub

Private Sub ch4_Click()
D1 = 150
End Sub

Private Sub cmd1_Click()
If cb1.Text = "Argentina" And cb2.Text = "Barco" Then
A = 1000
Subt = A1 + A
Subt = B1 + A
Subt = C1 + A
Subt = D1 + A
End If

If ch5 Then
Por = (Subt * 10) / 100
Total = Subt - Por
End If
If ch4 Then
Por = (Subt * 15) / 100
Total = Subt - Por
End If

If IsNumeric(txt2.Text) = True Then
lbl1.Caption = Val(txt2.Text) * Total
End If

If cb1.Text = "Argentina" And cb2.Text = "Avion" Then
B = 2000
lbl1.Caption = A1 + B
lbl1.Caption = B1 + B
lbl1.Caption = C1 + B
lbl1.Caption = D1 + B
End If

If ch5 Then
Por = (Subt * 10) / 100
Total = Subt - Por
End If
If ch4 Then
Por = (Subt * 15) / 100
Total = Subt - Por
End If

If IsNumeric(txt2.Text) = True Then
lbl1.Caption = Val(txt2.Text) * Total
End If

If cb1.Text = "Argentina" And cb2.Text = "Micro" Then
C = 500
lbl1.Caption = A1 + C
lbl1.Caption = B1 + C
lbl1.Caption = C1 + C
lbl1.Caption = D1 + C
End If

If ch5 Then
Por = (Subt * 10) / 100
Total = Subt - Por
End If
If ch4 Then
Por = (Subt * 15) / 100
Total = Subt - Por
End If

If IsNumeric(txt2.Text) = True Then
lbl1.Caption = Val(txt2.Text) * Total
End If

If cb1.Text = "Brazil" And cb2.Text = "Avion" Then
D = 2500
lbl1.Caption = A1 + D
lbl1.Caption = B1 + D
lbl1.Caption = C1 + D
lbl1.Caption = D1 + D
End If

If ch5 Then
Por = (Subt * 10) / 100
Total = Subt - Por
End If
If ch4 Then
Por = (Subt * 15) / 100
Total = Subt - Por
End If

If IsNumeric(txt2.Text) = True Then
lbl1.Caption = Val(txt2.Text) * Total
End If

If cb1.Text = "Brazil" And cb2.Text = "Barco" Then
E = 1500
lbl1.Caption = A1 + E
lbl1.Caption = B1 + E
lbl1.Caption = C1 + E
lbl1.Caption = D1 + E
End If

If cb1.Text = "Brazil" And cb2.Text = "Micro" Then
F = 1200
lbl1.Caption = A1 + F
lbl1.Caption = B1 + F
lbl1.Caption = C1 + F
lbl1.Caption = D1 + F
End If

If cb1.Text = "Egipto" And cb2.Text = "Avion" Then
G = 3000
lbl1.Caption = A1 + G
lbl1.Caption = B1 + G
lbl1.Caption = C1 + G
lbl1.Caption = D1 + G
End If

If cb1.Text = "Egipto" And cb2.Text = "Barco" Then
H = 2000
lbl1.Caption = A1 + H
lbl1.Caption = B1 + H
lbl1.Caption = C1 + H
lbl1.Caption = D1 + H
End If

If cb1.Text = "Egipto" And cb2.Text = "Micro" Then
MsgBox ("Agarra un Mapa BOBO")
End If

End Sub

