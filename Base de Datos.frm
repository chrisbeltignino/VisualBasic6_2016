VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   465
      Left            =   2430
      TabIndex        =   18
      Top             =   3825
      Width           =   1185
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Ultimo Registro"
      Height          =   735
      Left            =   5400
      TabIndex        =   16
      Top             =   5355
      Width           =   1185
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Siguiente Registro"
      Height          =   735
      Left            =   3780
      TabIndex        =   15
      Top             =   5310
      Width           =   1185
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Anterior Registro"
      Height          =   735
      Left            =   2250
      TabIndex        =   14
      Top             =   5310
      Width           =   1185
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Primer Registro"
      Height          =   735
      Left            =   720
      TabIndex        =   13
      Top             =   5265
      Width           =   1185
   End
   Begin VB.TextBox Text4 
      Height          =   465
      Left            =   2430
      TabIndex        =   12
      Top             =   3060
      Width           =   1185
   End
   Begin VB.TextBox Text3 
      Height          =   465
      Left            =   2430
      TabIndex        =   11
      Top             =   2295
      Width           =   1185
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Consultar"
      Height          =   465
      Left            =   4815
      TabIndex        =   8
      Top             =   4410
      Width           =   1185
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   465
      Left            =   4770
      TabIndex        =   7
      Top             =   3015
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Height          =   465
      Left            =   4770
      TabIndex        =   6
      Top             =   2205
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Bajo"
      Height          =   465
      Left            =   4770
      TabIndex        =   5
      Top             =   1485
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Alto"
      Height          =   465
      Left            =   4770
      TabIndex        =   4
      Top             =   765
      Width           =   1185
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   660
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4410
      Width           =   2130
   End
   Begin VB.TextBox Text2 
      Height          =   465
      Left            =   2430
      TabIndex        =   1
      Top             =   1530
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Left            =   2430
      TabIndex        =   0
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label Label5 
      Caption         =   "Sueldo"
      Height          =   465
      Left            =   765
      TabIndex        =   17
      Top             =   3870
      Width           =   1185
   End
   Begin VB.Label Label4 
      Caption         =   "Teléfono"
      Height          =   465
      Left            =   720
      TabIndex        =   10
      Top             =   3060
      Width           =   1185
   End
   Begin VB.Label Label3 
      Caption         =   "Dirección"
      Height          =   465
      Left            =   720
      TabIndex        =   9
      Top             =   2295
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "DNI"
      Height          =   465
      Left            =   720
      TabIndex        =   3
      Top             =   1530
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   465
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command5_Click()

End Sub
