VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FrSemaforo 
   Caption         =   "Semaforo"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSeguir 
      Caption         =   "Seguir"
      Height          =   465
      Left            =   2970
      TabIndex        =   5
      Top             =   1350
      Width           =   1230
   End
   Begin VB.CommandButton cmdParar 
      Caption         =   "Parar"
      Height          =   465
      Left            =   2970
      TabIndex        =   4
      Top             =   630
      Width           =   1230
   End
   Begin VB.CommandButton cmdVerde 
      Caption         =   "Verde"
      Height          =   465
      Left            =   1935
      TabIndex        =   2
      Top             =   2115
      Width           =   690
   End
   Begin VB.CommandButton cmdAmarillo 
      Caption         =   "Amarillo"
      Height          =   465
      Left            =   1935
      TabIndex        =   1
      Top             =   1395
      Width           =   690
   End
   Begin VB.CommandButton cmdRojo 
      Caption         =   "Rojo"
      Height          =   510
      Left            =   1935
      TabIndex        =   0
      Top             =   585
      Width           =   690
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4725
      Top             =   1845
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   2355
      Left            =   2745
      TabIndex        =   3
      Top             =   3690
      Width           =   3705
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6535
      _cy             =   4154
   End
   Begin VB.Shape ShAmarillo 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   510
      Left            =   765
      Top             =   1305
      Width           =   465
   End
   Begin VB.Shape ShVerde 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   510
      Left            =   765
      Top             =   2025
      Width           =   465
   End
   Begin VB.Shape ShRojo 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   510
      Left            =   765
      Top             =   540
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   2535
      Left            =   540
      Top             =   315
      Width           =   870
   End
End
Attribute VB_Name = "FrSemaforo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim T As Integer
Private Declare Sub PortOut Lib "IO.DLL" (ByVal Port As Integer, ByVal Data As Byte)


Private Sub cmdAmarillo_Click()
T = 8000
Timer1.Interval = 1000
PortOut &H378, 24
ShRojo.BackColor = vbBlack
ShAmarillo.BackColor = vbYellow
ShVerde.BackColor = vbBlack
End Sub
Private Sub cmdRojo_Click()
T = 3000
Timer1.Interval = 1000
ShRojo.BackColor = vbRed
ShAmarillo.BackColor = vbBlack
ShVerde.BackColor = vbBlack
End Sub
Private Sub cmdVerde_Click()
T = 11000
Timer1.Interval = 1000
ShRojo.BackColor = vbBlack
ShAmarillo.BackColor = vbBlack
ShVerde.BackColor = vbGreen
End Sub

Private Sub cmdParar_Click()
If Timer1.Enabled = True Then
Timer1.Enabled = False
End If
PortOut &H378, 0
End Sub

Private Sub cmdSeguir_Click()
If Timer1.Enabled = False Then
Timer1.Enabled = True
End If
PortOut &H378, 24
End Sub

Private Sub Form_Load()
ShRojo.BackColor = vbBlack
ShVerde.BackColor = vbBlack
ShAmarillo.BackColor = vbBlack
End Sub

Private Sub Timer1_Timer()
PortOut &H378, 24
T = T + Val(Timer1.Interval)
If T = 1000 Then
Me.WindowsMediaPlayer1.URL = "C:\Users\Public\Music\Sample Music\Kalimba.mp3"
ShRojo.BackColor = vbRed
End If
If T = 11000 Then
Me.WindowsMediaPlayer1.URL = "C:\Users\Public\Music\Sample Music\Sleep Away.mp3"
ShAmarillo.BackColor = vbYellow
ShRojo.BackColor = vbBlack
End If
If T = 15000 Then
Me.WindowsMediaPlayer1.URL = "C:\Users\Public\Music\Sample Music\Maid with the Flaxen Hair.mp3"
ShVerde.BackColor = vbGreen
ShAmarillo.BackColor = vbBlack
End If
If T = 25000 Then
Me.WindowsMediaPlayer1.URL = "C:\Users\Public\Music\Sample Music\Sleep Away.mp3"
ShAmarillo.BackColor = vbYellow
ShVerde.BackColor = vbBlack
End If
If T = 30000 Then
Me.WindowsMediaPlayer1.URL = "C:\Users\Public\Music\Sample Music\Kalimba.mp3"
ShRojo.BackColor = vbRed
ShAmarillo.BackColor = vbBlack
T = 1000
Timer1.Interval = 1000
End If

End Sub

