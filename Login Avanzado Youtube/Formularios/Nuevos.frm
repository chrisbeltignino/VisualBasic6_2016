VERSION 5.00
Begin VB.Form Nuevos 
   BorderStyle     =   0  'None
   Caption         =   "Nuevos"
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TextContrasena 
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox TextContrasena2 
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ComboBox Combotipo 
      Height          =   315
      ItemData        =   "Nuevos.frx":0000
      Left            =   2160
      List            =   "Nuevos.frx":000A
      TabIndex        =   8
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crear Cuenta"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox TextId 
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Confirmacion :"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuentas De Usuario"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Contraseña :"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario :"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo De Cuenta"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "Nuevos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
'si el campo usuario o contraseña esta vacio que me mande el cursor ahi
If TextId.Text = "" Then
    MsgBox " El campo Usuario no puede estar Vacio", "Alerta": TextId.SetFocus:
Exit Sub
If TextContrasena.Text = "" Then
    MsgBox " El campo Contraseña no puede estar Vacio", "Alerta": TextContrasena.SetFocus:
Exit Sub
If TextContrasena2.Text = "" Then
    MsgBox " El campo Confirmacion de Contraseña no puede estar Vacio", "Alerta": TextContrasena2.SetFocus:
Exit Sub
If Combotipo.Text = "" Then
    MsgBox " El campo Tipo no puede estar Vacio", "Alerta": Combotipo.SetFocus:
Exit Sub
'Mandar traer la base de datos
If TextContrasena.Text = TextContrasena2.Text Then
With Usuarios
        .Requery
        .AddNew
        !Usuario = TextId.Text
        !Contrasena = TextContrasena.Text
        !Confirmacion = TextContrasena2.Text
        !Tipo = Combotipo.Text
        .Update
       limpiar
MsgBox "Bienvenido"
Login.Show
Unload Me
End With
Else
MsgBox "Las Contraseñas No Coinsiden", vbInformation, "ERROR DE CONTRASEÑA"
TextContrasena.Text = "": TextContrasena2.Text = ""
End If
End If
End If
End If
End If
End Sub
Sub limpiar()
TextId.Text = ""
TextContrasena.Text = ""
TextContrasena2.Text = ""
Combotipo.Text = ""
TextId.SetFocus
End Sub

Private Sub Form_Load()
Usuarioss
End Sub
