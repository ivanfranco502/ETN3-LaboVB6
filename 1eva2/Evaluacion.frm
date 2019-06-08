VERSION 5.00
Begin VB.Form Evaluacion 
   Caption         =   "Numero Caprichoso"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton btnCaprichoso 
      Caption         =   "Verificar si es Caprichoso"
      Height          =   735
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtNumero 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblCaprichoso 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Resultado:"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblResultado 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ingrese un valor positivo:            Entre 0 - 128"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Evaluacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------CONSIGNA-------------------------
'Una aplicación deberá decir si un número entero
'positivo ingresado por el usuario es caprichoso.
'Un entero positivo es caprichoso cuando en binario
'tiene una cantidad impar de unos.
'-----------------------------------------------------
'Usuario: 12b01
'Alumno: Iván Franco
'Curso: 4º 2º

Private Sub Contador(ByVal CantRep As Byte)
If CantRep <> 0 Then            'Si la repeticion de 1 es diferente a 0 procede
    If CantRep Mod 2 = 0 Then      'Busca si es divisible por dos la repeticion
        lblCaprichoso.Caption = "No es Caprichoso" 'Imprime
    Else
        lblCaprichoso.Caption = "Es caprichoso"    'Imprime
    End If
Else
    lblCaprichoso.Caption = "No es Caprichoso" 'Si es 0 imprime
End If
End Sub
Private Function intBinario(ByRef Vector() As Byte) As String
Dim bytF As Byte
Dim Binario As String
Dim bytContador As Byte
Dim resto(7) As Integer
Dim intNumer As Byte
intNumer = txtNumero.Text

Select Case intNumer
    Case Is = 128
    Vector(0) = 1
    bytContador = bytContador + 1
    GoTo salto
    Case Is = 64
    Vector(1) = 1
    bytContador = bytContador + 1
    GoTo salto
    Case Is = 32
    Vector(2) = 1
    bytContador = bytContador + 1
    GoTo salto
    Case Is = 16
    Vector(3) = 1
    bytContador = bytContador + 1
    GoTo salto
    Case Is = 8
    Vector(4) = 1
    bytContador = bytContador + 1
    GoTo salto
    Case Is = 4
    Vector(5) = 1
    bytContador = bytContador + 1
    GoTo salto
    Case Is = 2
    Vector(6) = 1
    bytContador = bytContador + 1
    GoTo salto
    Case Is = 1
    Vector(7) = 1
    bytContador = bytContador + 1
    GoTo salto
End Select
    


If intNumer - 128 >= 0 Then
resto(0) = intNumer - 128
If resto(0) < 0 Then
resto(0) = 128
End If
Vector(0) = 1
bytContador = bytContador + 1
End If
If intNumer - resto(0) - 64 >= 0 Then
resto(1) = resto(0) - 64
If resto(1) < 0 Then
resto(1) = 64
End If
Vector(1) = 1
bytContador = bytContador + 1
End If
If intNumer - resto(0) - resto(1) - 32 >= 0 Then
resto(2) = resto(1) - 32
If resto(2) < 0 Then
resto(2) = 32
End If
Vector(2) = 1
bytContador = bytContador + 1
End If
If intNumer - resto(0) - resto(1) - resto(2) - 16 >= 0 Then
resto(3) = resto(2) - 16
If resto(3) < 0 Then
resto(3) = 16
End If
Vector(3) = 1
bytContador = bytContador + 1
End If
If intNumer - resto(0) - resto(1) - resto(2) - resto(3) - 8 >= 0 Then
resto(4) = resto(3) - 8
If resto(4) < 0 Then
resto(4) = 8
End If
Vector(4) = 1
bytContador = bytContador + 1
End If
If intNumer - resto(0) - resto(1) - resto(2) - resto(3) - resto(4) - 4 >= 0 Then
resto(5) = resto(4) - 4
If resto(5) < 0 Then
resto(5) = 4
End If
Vector(5) = 1
bytContador = bytContador + 1
End If
If intNumer - resto(0) - resto(1) - resto(2) - resto(3) - resto(4) - resto(5) - 2 >= 0 Then
resto(6) = resto(5) - 2
If resto(6) < 0 Then
resto(6) = 2
End If
Vector(6) = 1
bytContador = bytContador + 1
End If
If intNumer - resto(0) - resto(1) - resto(2) - resto(3) - resto(4) - resto(5) - resto(6) - 1 >= 0 Then
resto(7) = resto(6) - 1
If resto(7) < 0 Then
resto(7) = 1
End If
Vector(7) = 1
bytContador = bytContador + 1
End If
salto:
For bytF = 0 To 7
    Binario = Binario & Vector(bytF) 'Pega en un string los vectores con sus datos
Next bytF

Contador bytContador 'Se llama al procedimiento de Caprichosos

intBinario = Binario 'Se manda el dato que contendra la funcion
End Function
Private Sub Alerta()
Dim Incorrecto As VbMsgBoxResult
    Incorrecto = MsgBox("El valor ingresado no es correcto", vbOKOnly, "Valor Incorrecto")
    'Descripcion del error con dialogo en mensaje emergente
            If Incorrecto = vbOK Then 'Si se clikea OK
                txtNumero.Text = ""   'Se borra el contenido
            End If
End Sub
Private Sub btnCaprichoso_Click()
Dim BytVector(7) As Byte

If IsNumeric(txtNumero.Text) Then    'Valida si es numerico el valor ingresado
        If txtNumero.Text >= 0 Then  'Valida si es positivo el valor ingresado
            If txtNumero.Text <= 128 Then
            lblResultado.Caption = intBinario(BytVector()) ' Se imprime lo que da
                                                           ' la funcion
            Else
            Print "Me salio hasta aca, no logre hacerlo mas grande el numero"
            End If
        Else
        Alerta  'Sino, Muestra error
        End If
Else
    Alerta      'Muestra error
End If

End Sub

Private Sub btnSalir_Click()
End
End Sub
