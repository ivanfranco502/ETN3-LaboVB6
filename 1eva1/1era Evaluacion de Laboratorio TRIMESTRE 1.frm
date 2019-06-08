VERSION 5.00
Begin VB.Form eva1 
   Caption         =   "1era Evaluacion 1er Trimestre"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ScaleHeight     =   2955
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCalcular 
      Caption         =   "&Calcular resultado"
      Default         =   -1  'True
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtDado 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Resultado:"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblCasillero 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Casillero:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblResultado 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese los valores obtenidos en el dado."
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "eva1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Un juego consiste en que un jugador avance por una pista de 10 casilleros según
'los valores obtenidos con un dado.
'Avanza lo que indica el dado si se obtienen valores entre 1 y 3 inclusive, pero
'retrocede 1 casillero con un 4, 2 con un 5 y 3 con un 6.
'El usuario ingresa en la aplicación los valores obtenidos con 5 tiradas de dados
'en una sola cadena, el sistema deberá decir en qué casillero quedará.
'Si el casillero final es negativo mostrará un cero, o si supera la casilla de salida,
'mostrará un cartel que dice que sa ha pasado del límite.
'En caso de ser 10 deberá mostrar un cartel de ganador.
'Ejemplo: "12322" -> Ganador; "3332" -> Pasado; "12445" -> 0 por ser negativo.

'Alumno: Iván Franco
'Curso: 4º 2º
'Usuario: 12b01

Private Sub btnCalcular_Click()

Dim bytN As Byte                    'Se declara la variable de repeticion
Dim bytCantCar As Byte              'Se declara la variable que tendra los caracteres
Dim intJuego As Integer             'Se declara la variable que tendra los casilleros
Dim bytSuma As Byte                 'Se declara la variable que tendra la suma

bytSuma = 0                         'Se le da el valor a la suma
bytCantCar = Len(txtDado.Text)      'Se consiguen los caracteres de lo introducido
bytN = 0                            'Se le da el valor a la repeticion
intJuego = 0                        'Se le da el valor a los casilleros

For bytN = 1 To bytCantCar                  'Comienza la repeticion
    Select Case Mid(txtDado.Text, bytN, 1)  'Se extrae para analizar los caracteres
        Case Is = "1"                       'Condicion si el caracter es 1
            intJuego = intJuego + 1
            bytSuma = bytSuma + 1
        Case Is = "2"                       'Condicion si el caracter es 2
            intJuego = intJuego + 1
            bytSuma = bytSuma + 2
        Case Is = "3"                       'Condicion si el caracter es 3
            intJuego = intJuego + 1
            bytSuma = bytSuma + 3
        Case Is = "4"                       'Condicion si el caracter es 4
            intJuego = intJuego - 1
            bytSuma = bytSuma + 4
        Case Is = "5"                       'Condicion si el caracter es 5
            intJuego = intJuego - 2
            bytSuma = bytSuma + 5
        Case Is = "6"                       'Condicion si el caracter es 6
            intJuego = intJuego - 3
            bytSuma = bytSuma + 6
        Case Else                           'Si el caracter no es ninguno anterior
            lblResultado.Caption = "Ha ingresado en alguna tirada un valor equivocado"
            Exit For
    End Select
Next bytN                                   'Comienza de nuevo la repeticion
If bytSuma = 10 Then                        'Condicion si gana
    lblResultado.Caption = "GANADOR"
End If
If bytSuma < 0 Then                         'Condicion si es 0
    lblResultado.Caption = "PERDISTE"
End If
If bytSuma > 10 Then                        'Condicion si se pasa
    lblResultado.Caption = "PASADO"
End If
If intJuego < 0 Then                        'Condicion si casillero negativo
    lblCasillero.Caption = "0"
Else
    If intJuego = 5 Then                    'Condicion si caracter es 5
    lblCasillero.Caption = intJuego
    End If
    If intJuego < 5 Then                    'Condicion si caracter < 5
    lblCasillero.Caption = intJuego
    End If
    If intJuego > 5 Then                    'Condicion si caracter > 5
    lblCasillero.Caption = intJuego & "Pasado"
    End If
End If
'Fin del programa
End Sub

