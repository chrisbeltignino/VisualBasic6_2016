VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7 
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   300
   ClientWidth     =   16410
   LinkTopic       =   "Form7"
   Moveable        =   0   'False
   Picture         =   "consulta.frx":0000
   ScaleHeight     =   7110
   ScaleWidth      =   16410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   14760
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Volver al Menu"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   6360
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   0
      Top             =   5160
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   1296
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
      Connect         =   $"consulta.frx":58D3
      OLEDBString     =   $"consulta.frx":5964
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
      Bindings        =   "consulta.frx":59F5
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   9128
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
   Begin VB.Menu mnuconsultas 
      Caption         =   "Consultas"
      Index           =   0
      Begin VB.Menu mnunom 
         Caption         =   "Nombre y Apellido"
         Index           =   0
      End
      Begin VB.Menu mnucurso 
         Caption         =   "Nombre y Curso"
         Index           =   0
      End
      Begin VB.Menu mnumater 
         Caption         =   "Nombre y Materias Pendientes"
         Index           =   0
      End
      Begin VB.Menu mnuver 
         Caption         =   "Ver Todo"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form7.Hide
Form5.Show
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
DataGrid1.Columns(5).Width = 5000
Adodc1.Refresh
DataGrid1.Refresh
End Sub

Private Sub mnucurso_Click(Index As Integer)
Adodc1.RecordSource = "Select Nombre,Curso from alumnos"
Adodc1.Refresh
DataGrid1.Refresh
End Sub

Private Sub mnumater_Click(Index As Integer)
Adodc1.RecordSource = "Select Nombre,MateriasPendientes from alumnos"
Adodc1.Refresh
DataGrid1.Refresh
End Sub

Private Sub mnunom_Click(Index As Integer)
Adodc1.RecordSource = "Select Nombre,Apellido from alumnos"
Adodc1.Refresh
DataGrid1.Refresh
End Sub

Private Sub mnuver_Click(Index As Integer)
Adodc1.RecordSource = "Select Nombre,Apellido,DNI,Telefono,Curso,MateriaPendiente from alumnos"
Adodc1.Refresh
DataGrid1.Refresh
End Sub
