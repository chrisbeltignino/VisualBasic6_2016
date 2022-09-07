VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd9 
      Height          =   1005
      Left            =   3375
      TabIndex        =   8
      Top             =   3330
      Width           =   1005
   End
   Begin VB.CommandButton cmd8 
      Height          =   1005
      Left            =   2205
      TabIndex        =   7
      Top             =   3330
      Width           =   1005
   End
   Begin VB.CommandButton cm7 
      Height          =   1005
      Left            =   1035
      TabIndex        =   6
      Top             =   3330
      Width           =   1005
   End
   Begin VB.CommandButton cmd6 
      Height          =   1005
      Left            =   3375
      TabIndex        =   5
      Top             =   2160
      Width           =   1005
   End
   Begin VB.CommandButton cmd5 
      Height          =   1005
      Left            =   2205
      TabIndex        =   4
      Top             =   2160
      Width           =   1005
   End
   Begin VB.CommandButton cmd4 
      Height          =   1005
      Left            =   1035
      TabIndex        =   3
      Top             =   2160
      Width           =   1005
   End
   Begin VB.CommandButton cmd3 
      Height          =   1005
      Left            =   3375
      TabIndex        =   2
      Top             =   990
      Width           =   1005
   End
   Begin VB.CommandButton cmd2 
      Height          =   1005
      Left            =   2205
      TabIndex        =   1
      Top             =   990
      Width           =   1005
   End
   Begin VB.CommandButton cmd1 
      Height          =   1005
      Left            =   1035
      TabIndex        =   0
      Top             =   990
      Width           =   1005
   End
   Begin VB.Shape X9 
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   3285
      Shape           =   3  'Circle
      Top             =   3555
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape X8 
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   2115
      Shape           =   3  'Circle
      Top             =   3555
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape X7 
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   900
      Shape           =   3  'Circle
      Top             =   3555
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape O7 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   900
      Shape           =   3  'Circle
      Top             =   3555
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape O8 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   2115
      Shape           =   3  'Circle
      Top             =   3555
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape O9 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   3285
      Shape           =   3  'Circle
      Top             =   3555
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape O6 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   3285
      Shape           =   3  'Circle
      Top             =   2430
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape X6 
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   3285
      Shape           =   3  'Circle
      Top             =   2430
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape O4 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   945
      Shape           =   3  'Circle
      Top             =   2430
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape X4 
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   945
      Shape           =   3  'Circle
      Top             =   2430
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape O5 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   2115
      Shape           =   3  'Circle
      Top             =   2430
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape X5 
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   2115
      Shape           =   3  'Circle
      Top             =   2430
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape O1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   945
      Shape           =   3  'Circle
      Top             =   1215
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape O3 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   3285
      Shape           =   3  'Circle
      Top             =   1215
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape X3 
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   3285
      Shape           =   3  'Circle
      Top             =   1215
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape X2 
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   2115
      Shape           =   3  'Circle
      Top             =   1215
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape O2 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   2115
      Shape           =   3  'Circle
      Top             =   1215
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape X1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   945
      Shape           =   3  'Circle
      Top             =   1215
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim P As Byte

Private Sub cmd1_Click()
If P / 2 = Int(P / 2) Then
cmd1.Visible = False
O1.Visible = True
P = P + 1
Else
cmd1.Visible = False
X1.Visible = True
P = P + 1
End If

If O1.Visible = True And O2.Visible = True And O3.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O4.Visible = True And O5.Visible = True And O6.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O7.Visible = True And O8.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O4.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O2.Visible = True And O5.Visible = True And O8.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O6.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O5.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O5.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
End Sub

Private Sub cmd2_Click()
If P / 2 = Int(P / 2) Then
cmd2.Visible = False
O2.Visible = True
P = P + 1
Else
cmd2.Visible = False
X2.Visible = True
P = P + 1
End If

If O1.Visible = True And O2.Visible = True And O3.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O4.Visible = True And O5.Visible = True And O6.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O7.Visible = True And O8.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O4.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O2.Visible = True And O5.Visible = True And O8.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O6.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O5.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O5.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
End Sub

Private Sub cmd3_Click()
If P / 2 = Int(P / 2) Then
cmd3.Visible = False
O3.Visible = True
P = P + 1
Else
cmd3.Visible = False
X3.Visible = True
P = P + 1
End If

If O1.Visible = True And O2.Visible = True And O3.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O4.Visible = True And O5.Visible = True And O6.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O7.Visible = True And O8.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O4.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O2.Visible = True And O5.Visible = True And O8.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O6.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O5.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O5.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
End Sub

Private Sub cmd4_Click()
If P / 2 = Int(P / 2) Then
cmd4.Visible = False
O4.Visible = True
P = P + 1
Else
cmd4.Visible = False
X4.Visible = True
P = P + 1
End If

If O1.Visible = True And O2.Visible = True And O3.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O4.Visible = True And O5.Visible = True And O6.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O7.Visible = True And O8.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O4.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O2.Visible = True And O5.Visible = True And O8.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O6.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O5.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O5.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
End Sub

Private Sub cmd5_Click()
If P / 2 = Int(P / 2) Then
cmd5.Visible = False
O5.Visible = True
P = P + 1
Else
cmd5.Visible = False
X5.Visible = True
P = P + 1
End If

If O1.Visible = True And O2.Visible = True And O3.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O4.Visible = True And O5.Visible = True And O6.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O7.Visible = True And O8.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O4.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O2.Visible = True And O5.Visible = True And O8.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O6.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O5.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O5.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
    
If X1.Visible = True And X2.Visible = True And O3.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X4.Visible = True And X5.Visible = True And O6.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X7.Visible = True And X8.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X1.Visible = True And X4.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X2.Visible = True And X5.Visible = True And O8.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X3.Visible = True And X6.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X1.Visible = True And X5.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X3.Visible = True And X5.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Negro"
End Sub

Private Sub cmd6_Click()
If P / 2 = Int(P / 2) Then
cmd6.Visible = False
O6.Visible = True
P = P + 1
Else
cmd6.Visible = False
X6.Visible = True
P = P + 1
End If

If O1.Visible = True And O2.Visible = True And O3.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O4.Visible = True And O5.Visible = True And O6.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O7.Visible = True And O8.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O4.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O2.Visible = True And O5.Visible = True And O8.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O6.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O5.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O5.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
    
If X1.Visible = True And X2.Visible = True And O3.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X4.Visible = True And X5.Visible = True And O6.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X7.Visible = True And X8.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X1.Visible = True And X4.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X2.Visible = True And X5.Visible = True And O8.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X3.Visible = True And X6.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X1.Visible = True And X5.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X3.Visible = True And X5.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Negro"
End Sub

Private Sub cm7_Click()
If P / 2 = Int(P / 2) Then
cm7.Visible = False
O7.Visible = True
P = P + 1
Else
cm7.Visible = False
X7.Visible = True
P = P + 1
End If

If O1.Visible = True And O2.Visible = True And O3.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O4.Visible = True And O5.Visible = True And O6.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O7.Visible = True And O8.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O4.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O2.Visible = True And O5.Visible = True And O8.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O6.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O5.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O5.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
    
If X1.Visible = True And X2.Visible = True And O3.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X4.Visible = True And X5.Visible = True And O6.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X7.Visible = True And X8.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X1.Visible = True And X4.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X2.Visible = True And X5.Visible = True And O8.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X3.Visible = True And X6.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X1.Visible = True And X5.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X3.Visible = True And X5.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Negro"
End Sub

Private Sub cmd8_Click()
If P / 2 = Int(P / 2) Then
cmd8.Visible = False
O8.Visible = True
P = P + 1
Else
cmd8.Visible = False
X8.Visible = True
P = P + 1
End If

If O1.Visible = True And O2.Visible = True And O3.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O4.Visible = True And O5.Visible = True And O6.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O7.Visible = True And O8.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O4.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O2.Visible = True And O5.Visible = True And O8.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O6.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O5.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O5.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
    
If X1.Visible = True And X2.Visible = True And O3.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X4.Visible = True And X5.Visible = True And O6.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X7.Visible = True And X8.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X1.Visible = True And X4.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X2.Visible = True And X5.Visible = True And O8.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X3.Visible = True And X6.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X1.Visible = True And X5.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X3.Visible = True And X5.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Negro"
End Sub

Private Sub cmd9_Click()
If P / 2 = Int(P / 2) Then
cmd9.Visible = False
O9.Visible = True
P = P + 1
Else
cmd9.Visible = False
X9.Visible = True
P = P + 1
End If

If O1.Visible = True And O2.Visible = True And O3.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O4.Visible = True And O5.Visible = True And O6.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O7.Visible = True And O8.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O4.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O2.Visible = True And O5.Visible = True And O8.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O6.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O1.Visible = True And O5.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Rojo"
If O3.Visible = True And O5.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Rojo"

If X1.Visible = True And X2.Visible = True And O3.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X4.Visible = True And X5.Visible = True And O6.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X7.Visible = True And X8.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X1.Visible = True And X4.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X2.Visible = True And X5.Visible = True And O8.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X3.Visible = True And X6.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X1.Visible = True And X5.Visible = True And O9.Visible = True Then
    MsgBox "Gana el Jugador Negro"
If X3.Visible = True And X5.Visible = True And O7.Visible = True Then
    MsgBox "Gana el Jugador Negro"

End Sub
