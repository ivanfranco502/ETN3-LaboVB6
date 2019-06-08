VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSalir 
      Caption         =   "X"
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton BtnOrdenar 
      Caption         =   "Ordenar en Espiral"
      Height          =   495
      Left            =   960
      TabIndex        =   18
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtEspiral 
      Height          =   375
      Index           =   8
      Left            =   1800
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtEspiral 
      Height          =   375
      Index           =   7
      Left            =   1080
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtEspiral 
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtEspiral 
      Height          =   375
      Index           =   5
      Left            =   1800
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtEspiral 
      Height          =   375
      Index           =   4
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtEspiral 
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtEspiral 
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtEspiral 
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtEspiral 
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblEspiral 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   8
      Left            =   4200
      TabIndex        =   17
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblEspiral 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   7
      Left            =   3480
      TabIndex        =   16
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblEspiral 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   6
      Left            =   2760
      TabIndex        =   15
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblEspiral 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   5
      Left            =   4200
      TabIndex        =   14
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblEspiral 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   3480
      TabIndex        =   13
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblEspiral 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   12
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblEspiral 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblEspiral 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblEspiral 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub swap(ByRef x As Variant, y As Variant)
Dim z As Variant
z = x
x = y
y = z
End Sub
Private Sub ordenar(ByRef matriz() As Byte)
Dim byta As Byte
Dim bytb As Byte
Dim fila1 As Byte
Dim fila2 As Byte
Dim col1 As Byte
Dim col2 As Byte

For byta = 0 To 8
    fila1 = byta \ 3
    col1 = byta Mod 3
    For bytb = 0 To 8
        fila2 = bytb \ 3
        col2 = bytb Mod 3
        If matriz(fila1, col1) < matriz(fila2, col2) Then
            swap matriz(fila1, col1), matriz(fila2, col2)
        End If
    Next bytb
Next byta
Dim cuatro As Byte
Dim cinco As Byte
Dim seis As Byte
Dim ocho As Byte
Dim nueve As Byte

cuatro = matriz(2, 1)
cinco = matriz(2, 2)
seis = matriz(1, 0)
ocho = matriz(1, 2)
nueve = matriz(1, 1)

matriz(1, 0) = cuatro
matriz(1, 1) = cinco
matriz(1, 2) = seis
matriz(2, 1) = ocho
matriz(2, 2) = nueve

End Sub

Private Sub Espiral(ByRef matriz() As Byte)
Dim bytc As Byte
Dim bytf As Byte

Do: DoEvents
    For bytc = 0 To 2
    DoEvents
        Print matriz(bytf, bytc)
    Next bytc
    bytc = 2
    For bytf = 1 To 2
    DoEvents
        Print matriz(bytf, bytc)
    Next bytf
    bytf = 2
    Do
    DoEvents
        bytc = bytc - 1
        Print matriz(bytf, bytc)
    Loop Until bytc = 0
    Do
    DoEvents
        bytf = bytf - 1
        Print matriz(bytf, bytc)
    Loop Until bytf = 1
Print matriz(bytf, bytc + 1)
Loop Until (bytc = 0 And bytf = 1)
        
End Sub
Private Sub BtnOrdenar_Click()
Dim bytn As Byte
Dim pos As Byte
Dim Fila As Byte
Dim col As Byte
Dim matriz(2, 2) As Byte

For bytn = 0 To 8
    pos = bytn
    Fila = pos \ 3
    col = pos Mod 3
    matriz(Fila, col) = txtEspiral(bytn).Text
Next bytn
Espiral matriz()
ordenar matriz()
For bytn = 0 To 8
    pos = bytn
    Fila = pos \ 3
    col = pos Mod 3
    lblEspiral(bytn).Caption = matriz(Fila, col)
Next bytn
End Sub

Private Sub btnSalir_Click()
End
End Sub
