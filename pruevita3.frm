VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "."
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form3"
   ScaleHeight     =   6090
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Consulta"
      Height          =   510
      Left            =   6795
      TabIndex        =   18
      Top             =   4770
      Width           =   1185
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Formulario"
      Height          =   510
      Left            =   6705
      TabIndex        =   17
      Top             =   3960
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Alta"
      Height          =   510
      Left            =   7875
      TabIndex        =   16
      Top             =   3060
      Width           =   1185
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Usuario\Desktop\escfut.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Horarios"
      Top             =   4050
      Width           =   1860
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar Horarios"
      Height          =   510
      Left            =   7875
      TabIndex        =   15
      Top             =   1890
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar Horarios"
      Height          =   510
      Left            =   7830
      TabIndex        =   14
      Top             =   945
      Width           =   1185
   End
   Begin VB.TextBox Text7 
      DataField       =   "Domingo"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   5535
      TabIndex        =   6
      Top             =   2250
      Width           =   1185
   End
   Begin VB.TextBox Text6 
      DataField       =   "Sabado"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   5535
      TabIndex        =   5
      Top             =   1350
      Width           =   1185
   End
   Begin VB.TextBox Text5 
      DataField       =   "Viernes"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   5535
      TabIndex        =   4
      Top             =   405
      Width           =   1185
   End
   Begin VB.TextBox Text4 
      DataField       =   "Jueves"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   1980
      TabIndex        =   3
      Top             =   3015
      Width           =   1185
   End
   Begin VB.TextBox Text3 
      DataField       =   "Miercoles"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   1980
      TabIndex        =   2
      Top             =   2160
      Width           =   1185
   End
   Begin VB.TextBox Text2 
      DataField       =   "Martes"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   1980
      TabIndex        =   1
      Top             =   1305
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      DataField       =   "Lunes"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   1980
      TabIndex        =   0
      Top             =   450
      Width           =   1185
   End
   Begin VB.Label Label7 
      Caption         =   "Domingo"
      Height          =   510
      Left            =   3870
      TabIndex        =   13
      Top             =   2295
      Width           =   1185
   End
   Begin VB.Label Label6 
      Caption         =   "Sabado"
      Height          =   510
      Left            =   3780
      TabIndex        =   12
      Top             =   1395
      Width           =   1185
   End
   Begin VB.Label Label5 
      Caption         =   "Viernes"
      Height          =   510
      Left            =   3735
      TabIndex        =   11
      Top             =   495
      Width           =   1185
   End
   Begin VB.Label Label4 
      Caption         =   "Jueves"
      Height          =   510
      Left            =   405
      TabIndex        =   10
      Top             =   3060
      Width           =   1185
   End
   Begin VB.Label Label3 
      Caption         =   "Miercoles "
      Height          =   510
      Left            =   405
      TabIndex        =   9
      Top             =   2160
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "Martes"
      Height          =   510
      Left            =   450
      TabIndex        =   8
      Top             =   1305
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Lunes"
      Height          =   510
      Left            =   315
      TabIndex        =   7
      Top             =   450
      Width           =   1185
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" And Text5.Text = "" And Text6.Text = "" And Text7.Text = "" Then
MsgBox "Pone algo crack", vbOKOnly, "No se puede"
Else
Data1.Recordset.Update
Data1.Enabled = True
End If
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
Data1.Recordset.Edit
Data1.Enabled = True
End Sub

Private Sub Command3_Click()
Data1.Recordset.AddNew
End Sub

Private Sub Command4_Click()
Form3.Hide
Form1.Show
End Sub

Private Sub Command5_Click()
Form3.Hide
Form2.Show
End Sub
