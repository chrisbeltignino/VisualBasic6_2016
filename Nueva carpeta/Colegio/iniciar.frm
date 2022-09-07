VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10305
   LinkTopic       =   "Form2"
   Moveable        =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000E&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   4800
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   9480
      Top             =   5520
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000E&
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Iniciar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      MaskColor       =   &H80000014&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   8040
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "iniciar.frx":0000
      Left            =   8040
      List            =   "iniciar.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   5760
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   6615
      Left            =   -480
      Picture         =   "iniciar.frx":001F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Timer1.Enabled = True

End Sub

Private Sub Command2_Click()
Form3.Hide
Form2.Show
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
ProgressBar1.Value = 0
Label3.Caption = 0
Timer1.Enabled = False
Text1.Text = ""
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1 + 1
Label3.Caption = ProgressBar1.Value & "%"

If ProgressBar1.Value = 99 Then
    Timer1.Enabled = False
    ProgressBar1.Value = 0
    Label3.Caption = 0

If Combo1.Text = "Admin" And Text1.Text = "007" Then
    Form3.Hide
    Form5.Show
Else
If Combo1.Text = "Invitado" And Text1.Text = "321" Then
    Form3.Hide
    Form4.Show
Else
If Combo1.Text = "" Then
    MsgBox ("Ingrese un Usuario")
Else
If Text1.Text = "" Then
    MsgBox ("Ingrese la Contraseña")
Else
    MsgBox ("Contraseña Incorrecta")
End If
End If
End If
End If
End If
End Sub
