VERSION 5.00
Begin VB.Form Amigos 
   Caption         =   "Numeros Amigos"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "X"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton btnAmigos 
      Caption         =   "Click para ver los numeros amigos"
      Default         =   -1  'True
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
End
Attribute VB_Name = "Amigos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function intSumaDiv(ByVal intN As Integer) As Integer
Dim intDivisor As Integer, intSuma As Integer
intDivisor = 1
For intDivisor = 1 To Fix(intN / 2)
    If (intN Mod intDivisor = 0) Then intSuma = intSuma + intDivisor
Next intDivisor
intSumaDiv = intSuma
End Function

Private Sub btnAmigos_Click()
Dim intA As Integer, intB As Integer, bytCont As Byte
intA = 1
Do: DoEvents
    intB = intSumaDiv(intA)
        If (intSumaDiv(intB) = intA) And intA <> intB Then
            bytCont = bytCont + 1
        Print intA
        End If
    intA = intA + 1
Loop While Not bytCont = 6


End Sub

Private Sub btnSalir_Click()
End
End Sub
