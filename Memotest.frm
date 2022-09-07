VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   Picture         =   "Memotest.frx":0000
   ScaleHeight     =   6120
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   15
      Left            =   4185
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   15
      Top             =   4185
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   14
      Left            =   3105
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   14
      Top             =   4230
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   13
      Left            =   1980
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   13
      Top             =   4230
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   12
      Left            =   855
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   12
      Top             =   4275
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   11
      Left            =   4185
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   11
      Top             =   2925
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   10
      Left            =   3060
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   10
      Top             =   3015
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   9
      Left            =   1935
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   9
      Top             =   2925
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   8
      Left            =   810
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   8
      Top             =   3060
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   7
      Left            =   4140
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   7
      Top             =   1710
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   6
      Left            =   3015
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   6
      Top             =   1755
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   5
      Left            =   1935
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   5
      Top             =   1710
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   4
      Left            =   810
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   4
      Top             =   1755
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   3
      Left            =   4140
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   3
      Top             =   495
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   2
      Left            =   3015
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   2
      Top             =   495
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   1
      Left            =   1935
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   1
      Top             =   495
      Width           =   1005
   End
   Begin VB.PictureBox Imagen 
      Height          =   1095
      Index           =   0
      Left            =   810
      ScaleHeight     =   1035
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   495
      Width           =   1005
   End
   Begin VB.Label lbl1 
      Caption         =   "Fallos:"
      Height          =   465
      Left            =   3150
      TabIndex        =   17
      Top             =   5535
      Width           =   1185
   End
   Begin VB.Label lbl 
      Caption         =   "Aciertos:"
      Height          =   420
      Left            =   1215
      TabIndex        =   16
      Top             =   5580
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NumeroA As Integer
Dim NumeroF As Integer
Dim NumeroP As Integer
Dim Img1 As Integer
Dim Img2 As Integer
Dim Borrar As Boolean
Dim Ubi1(5, 5) As Integer
Dim Ubi2(5, 5) As Integer
Dim Memoria(4, 4) As Integer
Dim Aciertos(4, 4) As Boolean
Dim l As Integer

Private Sub Form_Load()
    Call EstablecerImagenes
    Call EstablecerMatriz
    NumeroP = 0
End Sub

Private Sub EstablecerImagenes()
    Dim i As Integer
    For i = 0 To 15
        Imagen(i).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\Interr.jpg")
    Next
End Sub

Private Sub EstablecerMatriz()
    Memoria(0, 0) = 1
    Memoria(1, 0) = 5
    Memoria(2, 0) = 6
    Memoria(3, 0) = 3
    Memoria(0, 1) = 2
    Memoria(1, 1) = 4
    Memoria(2, 1) = 8
    Memoria(3, 1) = 7
    Memoria(0, 2) = 1
    Memoria(1, 2) = 6
    Memoria(2, 2) = 8
    Memoria(3, 2) = 2
    Memoria(0, 2) = 4
    Memoria(3, 1) = 7
    Memoria(3, 2) = 5
    Memoria(3, 3) = 3
End Sub

Private Sub Interrogante()
    Dim ContadorF As Integer
    ContadorF = 0
    Dim ContadorC As Integer
    ContadorC = 0
    Dim l As Integer
    For l = 0 To 15
        If (Not (Aciertos(ContadorF, ContadorC))) Then
            Imagen(l).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\Interr.jpg")
        Else
            Imagen(l).Enabled = False
        End If
        
        If (ContadorC = 3) Then
            ContadorC = -1
            ContadorF = ContadorF + 1
        End If
        ContadorC = ContadorC + 1
    Next
End Sub

Private Sub Imagen_Click(Index As Integer)
    Select Case Index
        Case 0
            Imagen(0).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\rojo.jpg")
            Call Datos(0, 0)
            Imagen(l).Enabled = False
        Case 1
            Imagen(1).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\verde.jpg")
            Call Datos(0, 1)
            Imagen(l).Enabled = False
        Case 2
            Imagen(2).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\rosa.jpg")
            Call Datos(0, 2)
            Imagen(2).Enabled = False
        Case 3
            Imagen(3).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\azul.jpg")
            Call Datos(0, 3)
            Imagen(3).Enabled = False
        Case 4
            Imagen(4).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\violeta.jpg")
            Call Datos(1, 0)
            Imagen(4).Enabled = False
        Case 5
            Imagen(5).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\amarillo.jpg")
            Call Datos(1, 1)
            Imagen(5).Enabled = False
        Case 6
            Imagen(6).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\gris.jpg")
            Call Datos(1, 2)
            Imagen(6).Enabled = False
        Case 7
            Imagen(7).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\naranja.jpg")
            Call Datos(1, 3)
            Imagen(7).Enabled = False
        Case 8
            Imagen(8).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\rojo.jpg")
            Call Datos(2, 0)
            Imagen(8).Enabled = False
        Case 9
            Imagen(9).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\rojo.jpg")
            Call Datos(2, 1)
            Imagen(9).Enabled = False
        Case 10
            Imagen(10).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\rojo.jpg")
            Call Datos(2, 2)
            Imagen(10).Enabled = False
        Case 11
            Imagen(11).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\rojo.jpg")
            Call Datos(2, 3)
            Imagen(11).Enabled = False
        Case 12
            Imagen(12).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\rojo.jpg")
            Call Datos(3, 0)
            Imagen(12).Enabled = False
        Case 13
            Imagen(13).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\rojo.jpg")
            Call Datos(3, 1)
            Imagen(13).Enabled = False
        Case 14
            Imagen(14).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\rojo.jpg")
            Call Datos(3, 2)
            Imagen(14).Enabled = False
        Case 15
            Imagen(15).Picture = LoadPicture("C:\Users\Usuario\Desktop\Imagenes\rojo.jpg")
            Call Datos(3, 3)
            Imagen(15).Enabled = False
        End Select

End Sub

Private Sub Datos(Y As Integer, X As Integer)
    Select Case NumeroP
        Case 0
            Ubi1(0, 0) = Y
            Ubi1(0, 1) = X
            NumeroP = 1
            Img1 = Memoria(Y, X)
            Borrar = False
        Case 1
            Ubi2(0, 0) = Y
            Ubi2(0, 1) = X
            Img2 = Memoria(Y, X)
            NumeroP = 0
            Call Comparar(Img1, Img2)
            Borrar = True
    End Select
End Sub

Private Sub Comparar(Valor1 As Integer, Valor2 As Integer)
    If (Valor1 = Valor2) Then
        NumeroA = NumeroA + 1
        lbl.Caption = "Aciertos: " + CStr(NumeroA)
        Aciertos(Ubi1(0, 0), Ubi1(0, 1)) = True
        Aciertos(Ubi2(0, 0), Ubi2(0, 1)) = True
        MsgBox "Correcto"
    Else
        NumeroF = NumeroF + 1
        lbl1.Caption = "Fallos: " + CStr(NumeroF)
        Aciertos(Ubi1(0, 0), Ubi1(0, 1)) = False
        Aciertos(Ubi2(0, 0), Ubi2(0, 1)) = False
        MsgBox "Incorrecto"
    End If
    Img1 = 0
    Img2 = 0
End Sub
