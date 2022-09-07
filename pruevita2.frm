VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form2"
   ScaleHeight     =   6195
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Horario"
      Height          =   510
      Left            =   8640
      TabIndex        =   7
      Top             =   3600
      Width           =   1230
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Formulario"
      Height          =   510
      Left            =   6795
      TabIndex        =   6
      Top             =   3600
      Width           =   1230
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Todo de Alumnos"
      Height          =   465
      Left            =   7065
      TabIndex        =   5
      Top             =   4815
      Width           =   1230
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Categoria de Alumno"
      Height          =   465
      Left            =   5580
      TabIndex        =   4
      Top             =   4815
      Width           =   1230
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Dirección de Alumno"
      Height          =   465
      Left            =   4050
      TabIndex        =   3
      Top             =   4815
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Teléfono de Alumnos"
      Height          =   465
      Left            =   2475
      TabIndex        =   2
      Top             =   4815
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mail  de Alumnos"
      Height          =   465
      Left            =   990
      TabIndex        =   1
      Top             =   4815
      Width           =   1230
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "pruevita2.frx":0000
      Height          =   2760
      Left            =   405
      TabIndex        =   0
      Top             =   315
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   4868
      _Version        =   393216
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   1080
      Top             =   3780
      Width           =   3345
      _ExtentX        =   5900
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
      Connect         =   $"pruevita2.frx":0015
      OLEDBString     =   $"pruevita2.frx":013B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from escfut"
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
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Adodc1.RecordSource = "Select maila,nombrea,apellidoa from datos"
Adodc1.Refresh
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "Select telefonoa,nombrea,apellidoaTel from datos"
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
Adodc1.RecordSource = "Select direcciona,nombrea,apellidoa from datos"
Adodc1.Refresh
End Sub

Private Sub Command4_Click()
Adodc1.RecordSource = "Select categoriaa,nombrea,apellidoa from datos"
Adodc1.Refresh
End Sub

Private Sub Command5_Click()
Adodc1.RecordSource = "Select nombrea,apellidoa,DNIa,direcciona,categoriaa,cuotaa,telefonoa from datos"
Adodc1.Refresh
End Sub

Private Sub Command6_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Command7_Click()
Form2.Hide
Form3.Show
End Sub
