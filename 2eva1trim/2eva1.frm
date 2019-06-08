VERSION 5.00
Begin VB.Form eva21 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evaluacion de laboratorio"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnAzar 
      Caption         =   "Generar"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "X"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtPalabra 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label lblResultado 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "La palabra mas parecida:"
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label lblAzar 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   2880
      TabIndex        =   9
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblAzar 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblAzar 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblAzar 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblAzar 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Palabras Generadas:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese una palabra de 5 caracteres:"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "eva21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2da Evaluación del Pimer Trimestre
'Alumno: Iván Franco
'Curso: 4º 2º
'Cuenta de usuario: 12b01
'Consigna:
'El sistema deberá generar y mostrar 5 palabras de 5 caracteres al azar, tengan o no sentido.
'El usuario ingresa otra palabra de 5 caracteres y el sistema deberá indicar cual de sus
'palabras generadas es la más parecida a la que ingresó el usuario. - Una palabra se parece
'más a otra, cuantos más caracteres iguales tenga en la misma posición - .

Private Sub swap(ByRef varx As Variant, ByRef vary As Variant)
Dim varaux As Variant

varaux = varx                'Variable auxiliar, toma valor de Variable 1
varx = vary                 'Variable 1 toma valor de Variable 2
vary = varaux               'Variable 2 toma valor de Variable auxiliar

End Sub
Private Sub Ordenar(ByRef bytCon() As Byte, ByRef strV() As String)
Dim bytW As Byte
Dim bytT As Byte
Dim bytCuenta As Byte

    For bytW = 0 To 3                               'Sistema de repeticion
        For bytT = 1 To 4                           'busca el mayor
            If bytCon(bytT) > bytCon(bytW) Then     'si lo encuentra
                swap bytCon(bytT), bytCon(bytW)     'lo cambia
                swap strV(bytT), strV(bytW)
            End If
        Next bytT
    Next bytW
bytW = 0                                            'Reutilizacion de la variable
    For bytW = 0 To 4                  'busca si no hubo coincidencias
        If bytCon(bytW) = 0 Then
            bytCuenta = bytCuenta + 1
        End If
    Next bytW
If bytCuenta = 5 Then
lblResultado.Caption = "No hubo coincidencias"      'No hubo coincidencias
Else
lblResultado.Caption = strV(0)                      'Muestra el mayor ordenado
End If
End Sub
Private Sub Azar(ByRef strV() As String)
Dim bytN As Byte                        'Se crean dos contadores
Dim bytC As Byte

bytC = 0                                'Se iguala a cero los dos contadores
bytN = 0

    Do: DoEvents                        'Sistema de repeticion - indice por indice
        Do: DoEvents                    'Creacion al azar, letra por letra
            strV(bytC) = strV(bytC) & Chr(Fix((90 - 65 + 1) * Rnd + 65))
            bytN = bytN + 1
        Loop While Not (bytN > 4)       'Condicion, si se cumple sale
        bytN = 0                        'El contador se iguala a 0 para usarlo en el
        bytC = bytC + 1                 'proximo indice
    Loop While Not (bytC > 4)           'Condicion, si se cumple sale
    
bytC = 0                                'Reutilizacion de la variable, contador
    
    Do: DoEvents                        'Sistema de repeticion para imprimir
        lblAzar(bytC).Caption = strV(bytC) 'el vector en el vector de label
        bytC = bytC + 1
    Loop While Not (bytC > 4)           'Condicion, si se cumple sale
End Sub
Private Sub CompararCaracteres(ByRef strV() As String, bytFinal As Byte, strPal As String)
'Se reciben los datos
Dim bytN As Byte                    'Creacion de contador
Dim bytJ As Byte
Dim bytV(4) As Byte

bytN = 0                            'Contador = 0 para usarlo
    
    For bytN = 0 To bytFinal        'Sistema de repeticion, entra a cada indice
        For bytJ = 1 To 5
            If Mid(strPal, bytJ, 1) = Mid(strV(bytN), bytJ, 1) Then 'Compara letra por letra
                bytV(bytN) = bytV(bytN) + 1     'Suma letra acertada
            End If
        Next bytJ
    Next bytN
Ordenar bytV(), strV()
End Sub
Private Sub btnAzar_Click()
Dim StrVector(4) As String              'Se define el vector
Dim strParecido As String                   'Variable que tendra la palabra mas parecida
Dim strPalabraUsuario As String             'Variable que tendra la palabra del usuario
    If txtPalabra.Text = "" Then
        Print "Ingrese PALABRA"
        End
    End If
txtPalabra.Text = UCase(txtPalabra.Text)    'Se pone en Mayus palabra usuario
strPalabraUsuario = txtPalabra.Text         'Se almacena en la variable
Azar StrVector()



CompararCaracteres StrVector(), 4, strPalabraUsuario  'Se manda a procedimiento


End Sub
Private Sub btnSalir_Click()
End                         'Boton de salida
End Sub
Private Sub Form_Load()
Randomize Timer             'Carga siempre un Random diferente
End Sub


