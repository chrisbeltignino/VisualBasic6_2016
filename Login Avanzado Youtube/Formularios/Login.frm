VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Login 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Nuevo 
      Caption         =   "New"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "X"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Iniciar 
      Caption         =   "Iniciar"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5160
      Top             =   2040
   End
   Begin VB.ComboBox Combotipo 
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Login.frx":0000
      Left            =   2160
      List            =   "Login.frx":000A
      TabIndex        =   4
      Top             =   480
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
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
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox TextUsu 
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
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
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
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
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciar Como..."
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
'cerra formulario..... tambn podrias cambiarlo por "END"
Unload Me
End Sub

Private Sub Form_Load()
'Cargar BD Al formulario
Usuarioss
End Sub

Private Sub Iniciar_Click()
If TextUsu = "" Then
    MsgBox "Ingrese nombre De Usuario", vbInformation, "Aviso": TextUsu.SetFocus:
Exit Sub
If TextContrasena = "" Then
    MsgBox "Ingrese contraseña De Usuario", vbInformation, "Aviso": TextContrasena.SetFocus:
Exit Sub
If Combotipo = "" Then
    MsgBox "Ingrese contraseña De Usuario", vbInformation, "Aviso": Combotipo.SetFocus:
Exit Sub
With Usuarios
    .Requery 'actualizar tabla
    'Busqueda De Comparacion Con El Usuario
    .Find "Usuario='" & Trim(TextUsu.Text) & "'"
    If .EOF Then 'Si no se encuentra nada
    MsgBox "Usuario Incorrecto", vbInformation, "Aviso" 'msj de error
    TextUsu.Text = ""
    Exit Sub 'dejar de ejecutar
    Else
     If !Tipo = Trim(Combotipo.Text) Then 'pregunta si clave es correcta
    If !Contrasena = Trim(TextContrasena.Text) Then 'pregunta si clave es correcta
   ProgressBar1.Visible = True
   Timer1.Enabled = True
   TextUsu.Enabled = False
   TextContrasena.Enabled = False
   Combotipo.Enabled = False
    Else
    MsgBox "Contraseña Incorrecta", vbInformation, "Aviso"
    TextContrasena.Text = ""
    Exit Sub
        End If
        End If
        End If
 End With
End If
End If
End If
End Sub

Private Sub Nuevo_Click()
'Abrir formulario acceso y cerrar formulario Login
Acceso.Show
Unload Me
End Sub

Private Sub Timer1_Timer()
' el valor del progressbar es = al progressbar + 1
ProgressBar1.Value = ProgressBar1 + 1
'si el valor del progressbar es = a 1 entonces el label sera = a ( Cargando... )
    If ProgressBar1.Value = 1 Then
    Label4.Caption = "Cargando......."
    End If
    If ProgressBar1.Value = 20 Then
    Label4.Caption = "Comprobando Usuario"
    End If
    If ProgressBar1.Value = 40 Then
    Label4.Caption = "Comprobando Contraseña"
     End If
    If ProgressBar1.Value = 60 Then
    Label4.Caption = "El Sistema Esta Por Iniciar"
    End If
    If ProgressBar1.Value = 70 Then
    Label4.Caption = "Usuario Y Contraseña Correctos"
    End If
    If ProgressBar1.Value = 80 Then
    Label4.Caption = "Bienvenido-" + TextUsu.Text
     End If
    If ProgressBar1.Value = 100 Then

    Principal.Show
    Unload Me
    End If
End Sub
