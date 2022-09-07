VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16155
   LinkTopic       =   "Form6"
   Moveable        =   0   'False
   Picture         =   "alumns.frx":0000
   ScaleHeight     =   5475
   ScaleWidth      =   16155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Actualizar"
      Height          =   495
      Left            =   4080
      TabIndex        =   19
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Salir"
      Height          =   495
      Left            =   14640
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Volver al Menu"
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "MateriaPendientes"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   1965
      MaxLength       =   10000
      TabIndex        =   16
      Top             =   3000
      Width           =   1635
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   5640
      Top             =   3840
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"alumns.frx":58D3
      OLEDBString     =   $"alumns.frx":5964
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from alumnos"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "alumns.frx":59F5
      Height          =   3375
      Left            =   5640
      TabIndex        =   14
      Top             =   480
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5953
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      DataField       =   "Nombre"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   1965
      MaxLength       =   50
      TabIndex        =   8
      Top             =   600
      Width           =   1635
   End
   Begin VB.TextBox Text2 
      DataField       =   "Apellido"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   1965
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1080
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Alta"
      Height          =   375
      Left            =   4125
      TabIndex        =   6
      Top             =   720
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Baja"
      Height          =   375
      Left            =   4125
      TabIndex        =   5
      Top             =   1200
      Width           =   1005
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4125
      TabIndex        =   4
      Top             =   1680
      Width           =   1005
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4125
      TabIndex        =   3
      Top             =   2160
      Width           =   1005
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Chris\Desktop\Visual Basic\Nueva carpeta\Colegio\alumns.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "alumnos"
      Top             =   3840
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      DataField       =   "DNI"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   1965
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1560
      Width           =   1635
   End
   Begin VB.TextBox Text4 
      DataField       =   "Telefono"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   1965
      MaxLength       =   8
      TabIndex        =   1
      Top             =   2040
      Width           =   1635
   End
   Begin VB.TextBox Text5 
      DataField       =   "Curso"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   1965
      MaxLength       =   3
      TabIndex        =   0
      Top             =   2520
      Width           =   1635
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Materias Pendientes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   480
      TabIndex        =   15
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   770
      TabIndex        =   13
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   720
      TabIndex        =   12
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "DNI:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   960
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   690
      TabIndex        =   10
      Top             =   2100
      Width           =   1110
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Curso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   330
      Left            =   840
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A As String

Private Sub Command1_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Data1.Enabled = False
Data1.Recordset.AddNew

End Sub

Private Sub Command2_Click()
A = MsgBox("¿Esta seguro que deseea borrar esto?", vbYesNo, "Borrar")
If A = vbYes Then
Data1.Recordset.Delete
Data1.Refresh
Adodc1.Refresh
DataGrid1.Refresh
Else
Exit Sub
End If
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
MsgBox "Debe llenar todos los registros vacíos", vbOKOnly, "Error"
Else
Data1.Recordset.Update
Data1.Refresh
Adodc1.Refresh
DataGrid1.Refresh
Data1.Enabled = True
Command2.Enabled = True
Command3.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False

End If
End Sub

Private Sub Command4_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True

Command3.Enabled = True

Data1.Enabled = False
Data1.Recordset.Edit
Adodc1.Refresh
DataGrid1.Refresh

End Sub

Private Sub Command5_Click()
Form6.Hide
Form5.Show
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Command7_Click()
Adodc1.Refresh
End Sub

Private Sub Form_Load()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

Adodc1.Refresh
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97) And (KeyAscii < 122) Or (KeyAscii >= 65) And (KeyAscii < 90) Then
  KeyAscii = 0
End If
End Sub

