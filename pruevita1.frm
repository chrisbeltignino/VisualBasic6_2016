VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "Consultas"
      Height          =   510
      Left            =   7650
      TabIndex        =   36
      Top             =   6075
      Width           =   1185
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Horarios"
      Height          =   510
      Left            =   7560
      TabIndex        =   35
      Top             =   5265
      Width           =   1185
   End
   Begin VB.TextBox Text13 
      DataField       =   "categoriaa"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   5445
      TabIndex        =   31
      Top             =   6615
      Width           =   1185
   End
   Begin VB.TextBox Text12 
      DataField       =   "telefonoa"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   5445
      TabIndex        =   30
      Top             =   5715
      Width           =   1185
   End
   Begin VB.TextBox Text11 
      DataField       =   "maila"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   5445
      TabIndex        =   29
      Top             =   4815
      Width           =   1185
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Ultimo Registro"
      Height          =   510
      Left            =   1755
      TabIndex        =   28
      Top             =   6660
      Width           =   1185
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Anterior Registro"
      Height          =   510
      Left            =   1755
      TabIndex        =   27
      Top             =   5760
      Width           =   1185
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Siguiente Registro"
      Height          =   510
      Left            =   270
      TabIndex        =   26
      Top             =   6660
      Width           =   1185
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Primer Registro"
      Height          =   510
      Left            =   180
      TabIndex        =   25
      Top             =   5805
      Width           =   1185
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Consulta"
      Height          =   510
      Left            =   7560
      TabIndex        =   24
      Top             =   3960
      Width           =   1185
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   510
      Left            =   7560
      TabIndex        =   23
      Top             =   3015
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Height          =   510
      Left            =   7560
      TabIndex        =   22
      Top             =   2160
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Eliminar"
      Height          =   510
      Left            =   7560
      TabIndex        =   21
      Top             =   1215
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   510
      Left            =   7560
      TabIndex        =   20
      Top             =   360
      Width           =   1185
   End
   Begin VB.TextBox Text10 
      DataField       =   "cuotaa"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   5265
      TabIndex        =   9
      Top             =   3870
      Width           =   1185
   End
   Begin VB.TextBox Text9 
      DataField       =   "direcciona"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   5265
      TabIndex        =   8
      Top             =   2925
      Width           =   1185
   End
   Begin VB.TextBox Text8 
      DataField       =   "DNIa"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   5265
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1980
      Width           =   1185
   End
   Begin VB.TextBox Text7 
      DataField       =   "apellidoa"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   5265
      TabIndex        =   6
      Top             =   1125
      Width           =   1185
   End
   Begin VB.TextBox Text6 
      DataField       =   "nombrea"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   5265
      TabIndex        =   5
      Top             =   270
      Width           =   1185
   End
   Begin VB.TextBox Text5 
      DataField       =   "sueldop"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   1665
      TabIndex        =   4
      Top             =   3825
      Width           =   1185
   End
   Begin VB.TextBox Text4 
      DataField       =   "direccionp"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   1665
      TabIndex        =   3
      Top             =   2880
      Width           =   1185
   End
   Begin VB.TextBox Text3 
      DataField       =   "DNIp"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   1665
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1980
      Width           =   1185
   End
   Begin VB.TextBox Text2 
      DataField       =   "apellidop"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   1665
      TabIndex        =   1
      Top             =   1125
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      DataField       =   "nombrep"
      DataSource      =   "Data1"
      Height          =   510
      Left            =   1665
      TabIndex        =   0
      Top             =   270
      Width           =   1185
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Usuario\Desktop\escfut.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   225
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "datos"
      Top             =   4770
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "Categoria de Alumno"
      Height          =   510
      Left            =   3600
      TabIndex        =   34
      Top             =   6660
      Width           =   1185
   End
   Begin VB.Label Label12 
      Caption         =   "Teléfono de Alumno"
      Height          =   510
      Left            =   3555
      TabIndex        =   33
      Top             =   5760
      Width           =   1185
   End
   Begin VB.Label Label11 
      Caption         =   "Mail de Alumno"
      Height          =   510
      Left            =   3600
      TabIndex        =   32
      Top             =   4815
      Width           =   1185
   End
   Begin VB.Label Label10 
      Caption         =   "Cuota mensual "
      Height          =   510
      Left            =   3600
      TabIndex        =   19
      Top             =   3960
      Width           =   1185
   End
   Begin VB.Label Label9 
      Caption         =   "Direccion del Alumno"
      Height          =   510
      Left            =   3600
      TabIndex        =   18
      Top             =   2925
      Width           =   1185
   End
   Begin VB.Label Label8 
      Caption         =   "DNI Alumno"
      Height          =   510
      Left            =   3600
      TabIndex        =   17
      Top             =   2025
      Width           =   1185
   End
   Begin VB.Label Label7 
      Caption         =   "Apellido de Alumno"
      Height          =   510
      Left            =   3600
      TabIndex        =   16
      Top             =   1215
      Width           =   1185
   End
   Begin VB.Label Label6 
      Caption         =   "Nombre de Alumno"
      Height          =   510
      Left            =   3600
      TabIndex        =   15
      Top             =   225
      Width           =   1185
   End
   Begin VB.Label Label5 
      Caption         =   "Sueldo Mensual"
      Height          =   510
      Left            =   225
      TabIndex        =   14
      Top             =   3825
      Width           =   1185
   End
   Begin VB.Label Label4 
      Caption         =   "Direccion Profesor"
      Height          =   510
      Left            =   225
      TabIndex        =   13
      Top             =   2880
      Width           =   1185
   End
   Begin VB.Label Label3 
      Caption         =   "DNI Profesor"
      Height          =   510
      Left            =   225
      TabIndex        =   12
      Top             =   1980
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "Apellido de Profesor"
      Height          =   510
      Left            =   225
      TabIndex        =   11
      Top             =   1125
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de Profesor"
      Height          =   510
      Left            =   225
      TabIndex        =   10
      Top             =   270
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A As Integer

Private Sub Command1_Click()
Command2.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Data1.Recordset.AddNew
End Sub

Private Sub Command10_Click()
Form1.Hide
Form3.Show
End Sub

Private Sub Command11_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Command2_Click()
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" And Text5.Text = "" And Text6.Text = "" And Text7.Text = "" And Text8.Text = "" And Text9.Text = "" And Text10.Text = "" And Text11.Text = "" And Text12.Text And Text13.Text = "" Then
MsgBox "Pone algo crack", vbCritical, "Debes llenar el formulario para borrar"
Exit Sub
Else
A = MsgBox("¿ta´ lokito bo´ Quere borrar esto posta? ", vbYesNo, "Baja")
End If

If tuvi = vbYes Then
Data1.Recordset.Delete
Data1.Refresh
Else
Exit Sub
End If

End Sub

Private Sub Command3_Click()
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" And Text5.Text = "" And Text6.Text = "" And Text7.Text = "" And Text8.Text = "" And Text9.Text = "" And Text10.Text = "" And Text11.Text = "" And Text12.Text And Text13.Text = "" Then
MsgBox "Pone algo crack", vbOKOnly, "No se puede"
Else
Data1.Recordset.Update
Data1.Enabled = True
Command2.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
End If
End Sub

Private Sub Command4_Click()
Command1.Enabled = False
Command2.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Data1.Recordset.Edit
Data1.Enabled = True
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

Private Sub Text3_Change()

End Sub
