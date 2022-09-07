VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10890
   LinkTopic       =   "Form2"
   ScaleHeight     =   8355
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd6 
      Caption         =   "Direccion"
      Height          =   600
      Left            =   2700
      TabIndex        =   6
      Top             =   6255
      Width           =   1680
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "Sueldo"
      Height          =   600
      Left            =   4050
      TabIndex        =   5
      Top             =   5310
      Width           =   1680
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "Telefono"
      Height          =   600
      Left            =   1440
      TabIndex        =   4
      Top             =   5355
      Width           =   1680
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Todo"
      Height          =   600
      Left            =   5085
      TabIndex        =   3
      Top             =   4185
      Width           =   1680
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "DNI"
      Height          =   600
      Left            =   990
      TabIndex        =   2
      Top             =   4185
      Width           =   1680
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Nombre"
      Height          =   600
      Left            =   3060
      TabIndex        =   1
      Top             =   4185
      Width           =   1680
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Base de Datos.frx":0000
      Height          =   2535
      Left            =   900
      TabIndex        =   0
      Top             =   450
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   4471
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
            LCID            =   3082
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
            LCID            =   3082
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
      Height          =   690
      Left            =   900
      Top             =   3015
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   1217
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
      Connect         =   $"Base de Datos.frx":0015
      OLEDBString     =   $"Base de Datos.frx":0107
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   " select *  from Clientes"
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

Private Sub cmd1_Click()
Adodc1.RecordSource = "Select Nombre from Clientes"
Adodc1.Refresh
End Sub

Private Sub cmd3_Click()
Adodc1.RecordSource = "Select DNI from Clientes"
Adodc1.Refresh
End Sub

Private Sub cmd2_Click()
Adodc1.RecordSource = "Select Nombre,Dirección,Telefono,Sueldo,DNI from Clientes"
Adodc1.Refresh
End Sub

Private Sub cmd4_Click()
Adodc1.RecordSource = "Select Telefono from Clientes"
Adodc1.Refresh
End Sub

Private Sub cmd5_Click()
Adodc1.RecordSource = "Select Sueldo from Clientes"
Adodc1.Refresh
End Sub

Private Sub cmd6_Click()
Adodc1.RecordSource = "Select Dirección from Clientes"
Adodc1.Refresh
End Sub
