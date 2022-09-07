VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Chris\Desktop\Visual Basic\saco exam\empleados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "table"
      Top             =   5280
      Width           =   4455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Consultas"
      Height          =   975
      Left            =   9840
      TabIndex        =   15
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Salir"
      Height          =   975
      Left            =   7920
      TabIndex        =   14
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Guardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   8640
      TabIndex        =   13
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton Modificar 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   8640
      TabIndex        =   12
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton Baja 
      Caption         =   "Baja"
      Height          =   375
      Left            =   8640
      TabIndex        =   11
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton Alta 
      Caption         =   "Alta"
      Height          =   375
      Left            =   8640
      TabIndex        =   10
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      DataField       =   "mail"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   4320
      TabIndex        =   9
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      DataField       =   "tele"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   4320
      TabIndex        =   8
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      DataField       =   "dire"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   4320
      TabIndex        =   7
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombre"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   4320
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "codigo"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   4320
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Mail"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Telefono"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Direccion"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Alta_Click()
Data1.Recordset.AddNew
Baja.Enabled = False
Guardar.Enabled = True
Modificar.Enabled = False
End Sub

Private Sub Baja_Click()
Data1.Recordset.Delete
Data1.Refresh
End Sub

Private Sub Command6_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Guardar_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
MsgBox "Debes completar todos los datos", vbCritical, "Error"
Else
Data1.Recordset.Update
Modificar.Enabled = True
Baja.Enabled = True
Guardar.Enabled = False
End If
End Sub

Private Sub Modificar_Click()
Data1.Recordset.Edit
Alta.Enabled = False
Baja.Enabled = False
Modificar.Enabled = False
End Sub
