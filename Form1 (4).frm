VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Registros"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Alumno\Desktop\empleados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "empleados"
      Top             =   5520
      Width           =   2580
   End
   Begin VB.TextBox Text3 
      DataField       =   "horast"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   18
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdconsulta 
      Caption         =   "CONSULTAS"
      Height          =   495
      Left            =   4680
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdguardar 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdmod 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdbaja 
      Caption         =   "Baja"
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdalta 
      Caption         =   "Alta"
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      DataField       =   "fechan"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "totalc"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataField       =   "mail"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "horash"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombre"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "codigo"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Fecha de nacimiento"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Total a Cobrar"
      Height          =   495
      Left            =   480
      TabIndex        =   16
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Mail"
      Height          =   495
      Left            =   480
      TabIndex        =   15
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Precio por hora"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Horas Trabajadas"
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre y Apellido"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A As Integer
Dim B As String

Private Sub cmdalta_Click()

cmdbaja.Enabled = False
cmdguardar.Enabled = True
Data1.Enabled = False
Data1.Recordset.AddNew

End Sub

Private Sub cmdbaja_Click()

If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" And Text5.Text = "" And Text7.Text = "" Then
MsgBox "No se puede eliminar un registro vacío", vbOKOnly, "Error"
Exit Sub
Else
B = MsgBox("¿Esta seguro que deseea borrar esto?", vbYesNo, "Borrar")
If B = vbYes Then
Data1.Recordset.Delete
Data1.Refresh
Else
Exit Sub
End If

End If

End Sub

Private Sub cmdmod_Click()

cmdalta = False
cmdbaja = False
cmdguardar.Enabled = True
Data1.Enabled = False
Data1.Recordset.Edit

End Sub

Private Sub cmdguardar_Click()

If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" And Text5.Text = "" And Text7.Text = "" Then
MsgBox "No se puede guardar un registro vacío", vbOKOnly, "Error"
Else
Data1.Recordset.Update
Data1.Enabled = True
cmdbaja.Enabled = True
cmdguardar.Enabled = False
End If

End Sub

Private Sub cmdconsulta_Click()

Form2.Show

End Sub

Private Sub Text6_Change()

If Text3.Text = "" And Text4.Text = "" Then
MsgBox "No hay Total a Cobrar", vbOKOnly, "Error"
Else
A = Text3.Text * Text4.Text

End Sub

