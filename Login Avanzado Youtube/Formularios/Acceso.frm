VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Acceso 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Continuar 
      Caption         =   "Continuar"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3720
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ingresa Clave De Autorizacion "
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   2820
   End
End
Attribute VB_Name = "Acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Continuar_Click()
'si la caja de texto es = a 123 entonces, el timer sera = a verdadero y el progressbar sera visible
If Text1.Text = "123" Then
Timer1.Enabled = True: ProgressBar1.Visible = True
Else
'si es falso me mandara un msj.... caja de texto limpia y ubicar cursor en la misma
MsgBox "La clave no es correcta!...", vbInformation, "Aviso"
Text1.Text = ""
Text1.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
'el progressbar sera = al valor del progressbar + 1 " QUE ES COMO IRA CARGANDO"
ProgressBar1.Value = ProgressBar1 + 1
'si el valor del progressbar es = 100 entonces----- ire al fomulario nuevos
If ProgressBar1.Value = 100 Then
Nuevos.Show
Unload Me
End If
End Sub


