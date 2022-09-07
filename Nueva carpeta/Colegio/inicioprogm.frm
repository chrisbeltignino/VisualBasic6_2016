VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   2880
   End
   Begin VB.PictureBox ProgressBar1 
      Height          =   615
      Left            =   480
      ScaleHeight     =   555
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   3840
      Width           =   8055
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   4680
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   3600
      Left            =   2880
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      Picture         =   "inicioprogm.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
Label1.Caption = ProgressBar1.Value & "%"

If (ProgressBar1.Value = ProgressBar1.Max) Then
    Timer1.Enabled = False
    MsgBox ("Bienvenido al Sistema")
    Form1.Hide
    Form2.Show
End If

End Sub
