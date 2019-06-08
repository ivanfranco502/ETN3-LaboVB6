VERSION 5.00
Begin VB.Form TP2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Van's Primo Calculator"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPerfectos 
      BackColor       =   &H00808080&
      Caption         =   "Tres Numer&os Perfectos"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton btnPrimosEspecial 
      BackColor       =   &H00808080&
      Caption         =   "Li&stado de los Primeros N Primos"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox txtPrimos 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton btnSalir 
      BackColor       =   &H00808080&
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtListPrimos 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton btnListprimos 
      BackColor       =   &H00808080&
      Caption         =   "&Listado de Primos"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton btnPrimo 
      BackColor       =   &H00808080&
      Caption         =   "&Calcular"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtNumero 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.Line Line4 
      X1              =   7440
      X2              =   7440
      Y1              =   2400
      Y2              =   4680
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clickee y adquirira los primero 3 numeros primos perfectos"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   11
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Introduzca un valor para buscar los N primeros primos"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Line Line3 
      X1              =   480
      X2              =   11040
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   480
      Y1              =   0
      Y2              =   7680
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Introduzca un valor para saber si es primo o no."
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Introduzca un valor para buscar primos. 1 - Num"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   3360
      X2              =   3360
      Y1              =   0
      Y2              =   7680
   End
   Begin VB.Label lblResultado 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   1560
      Width           =   4695
   End
End
Attribute VB_Name = "TP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnListprimos_Click()
Dim bytNum As Byte, bytDivisor As Byte, intCont As Integer
bytNum = Val(txtListPrimos.Text)            'Se extrae el lapso de la busqueda
    Cls
    For bytNum = 1 To bytNum                'Se crea la busqueda de primos en el lapso
        intCont = 0                         'Se resetea el contador
            For bytDivisor = 1 To Fix(bytNum ^ 0.5)
                If bytNum Mod bytDivisor = 0 Then
                    intCont = intCont + 1   'Se le suma al contador los primos
                End If
            Next bytDivisor
                If intCont < 2 Then
                    Print bytNum      'Se imprimen los primos
                End If
    Next bytNum

End Sub

Private Sub btnPerfectos_Click()
Dim bytContPerf As Byte, intN As Integer, intDivisor As Integer
Dim intContDivis As Integer

    bytContPerf = 0
    intN = 1
    
    Cls
    Do: DoEvents
    intContDivis = 0
        For intDivisor = 1 To Fix(intN / 2)
                If intN Mod intDivisor = 0 Then
                    intContDivis = intContDivis + intDivisor
                End If
        Next intDivisor
                If intContDivis = intN Then
                    Print intN
                    bytContPerf = bytContPerf + 1
                End If
                intN = intN + 1
    Loop While Not (bytContPerf = 3)
End Sub

Private Sub btnPrimo_Click()
Dim bytDivisor As Byte, bytNum As Byte, intCont As Integer
    bytNum = Val(txtNumero.Text)       'Se extrae el numero ingresado
    intCont = 0                        'Se ingresa el primer valor al contador
    lblResultado.Caption = ""          'Se borra el Label de salida
            For bytDivisor = 1 To Fix(bytNum ^ 0.5)     'Se crea una repeticion para
                If bytNum Mod bytDivisor = 0 Then       'buscar primos
                    intCont = intCont + 1
                End If
            Next bytDivisor
                If intCont < 2 Then         'Si extrae el valor del contador
                    'Muestra resultado positivo
                        lblResultado.Caption = "El numero introducido es primo"
                Else
                    'Muestra resultado negativo
                        lblResultado.Caption = "El numero introducido es compuesto"
                End If
End Sub

Private Sub btnPrimosEspecial_Click()
Dim bytDivisor As Byte, bytNum As Byte, intCont As Integer, bytContPrimos As Byte
    bytNum = Val(txtNumero.Text)       'Se extrae el numero ingresado
    intCont = 0                        'Se ingresa el primer valor al contador
    bytContPrimos = 0
    lblResultado.Caption = ""          'Se borra el Label de salida
    bytNum = 1
    Cls
    Do
        intCont = 0
        For bytDivisor = 1 To Fix(bytNum ^ 0.5)
                If bytNum Mod bytDivisor = 0 Then
                    intCont = intCont + 1   'Se le suma al contador los primos
                End If
            Next bytDivisor
                If intCont < 2 Then
                    Print bytNum      'Se imprimen los primos
                    bytContPrimos = bytContPrimos + 1
                End If
            bytNum = bytNum + 1
    Loop While Not (bytContPrimos = txtPrimos.Text)
    
End Sub

Private Sub btnSalir_Click()
End
End Sub


