VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Ultimo registro"
      Height          =   465
      Left            =   4005
      TabIndex        =   18
      Top             =   5130
      Width           =   1005
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Siguiente registro"
      Height          =   465
      Left            =   2925
      TabIndex        =   17
      Top             =   5130
      Width           =   1005
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Anterior registro"
      Height          =   465
      Left            =   1845
      TabIndex        =   16
      Top             =   5130
      Width           =   1005
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Primer registro"
      Height          =   465
      Left            =   765
      TabIndex        =   15
      Top             =   5130
      Width           =   1005
   End
   Begin VB.TextBox Text5 
      DataField       =   "Sueldo"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   14
      Top             =   3420
      Width           =   1635
   End
   Begin VB.TextBox Text4 
      DataField       =   "Telefono"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   13
      Top             =   2970
      Width           =   1635
   End
   Begin VB.TextBox Text3 
      DataField       =   "Dirección"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   12
      Top             =   2520
      Width           =   1635
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Administrador\Mis documentos\Clientes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1260
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   4455
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Consulta"
      Height          =   465
      Left            =   5040
      TabIndex        =   8
      Top             =   3600
      Width           =   1140
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   2970
      Width           =   1005
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   2475
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Baja"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   1935
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Alta"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   1395
      Width           =   1005
   End
   Begin VB.TextBox Text2 
      DataField       =   "Dni"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2115
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      DataField       =   "Nombre"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1665
      Width           =   1635
   End
   Begin VB.Label Label5 
      Caption         =   "Sueldo:"
      Height          =   330
      Left            =   1035
      TabIndex        =   11
      Top             =   3465
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Telefono:"
      Height          =   330
      Left            =   1035
      TabIndex        =   10
      Top             =   3015
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Direccion:"
      Height          =   330
      Left            =   1035
      TabIndex        =   9
      Top             =   2565
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Dni:"
      Height          =   330
      Left            =   1035
      TabIndex        =   2
      Top             =   2160
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre:"
      Height          =   330
      Left            =   1035
      TabIndex        =   0
      Top             =   1710
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A As String

Private Sub Command1_Click()

Command2.Enabled = False
Command3.Enabled = True
Data1.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Data1.Recordset.AddNew


End Sub

Private Sub Command2_Click()
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" And Text5.Text = "" Then
MsgBox "No se puede eliminar un registro vacío", vbOKOnly, "Error"
Exit Sub
Else
A = MsgBox("¿Esta seguro que deseea borrar esto?", vbYesNo, "Borrar")
If A = vbYes Then
Data1.Recordset.Delete
Data1.Refresh
Else
Exit Sub
End If

End If
End Sub

Private Sub Command3_Click()
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" And Text5.Text = "" Then
MsgBox "No se puede guardar un registro vacío", vbOKOnly, "Error"
Else
Data1.Recordset.Update
Data1.Enabled = True
Command2.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command3.Enabled = False
End If
End Sub

Private Sub Command4_Click()
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command3.Enabled = True
Data1.Enabled = False
Data1.Recordset.Edit

End Sub

Private Sub Command6_Click()

Data1.Recordset.MoveFirst

End Sub

Private Sub Command7_Click()
If Data1.Recordset.BOF = True Then
Data1.Recordset.MoveNext
Else
Data1.Recordset.MovePrevious
End If
End Sub

Private Sub Command8_Click()
If Data1.Recordset.EOF = True Then
Data1.Recordset.MovePrevious
Else
Data1.Recordset.MoveNext
End If
End Sub

Private Sub Command9_Click()

Data1.Recordset.MoveLast
End Sub

