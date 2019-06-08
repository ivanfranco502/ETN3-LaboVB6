VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnEspiral 
      Caption         =   "Espiral"
      Height          =   615
      Left            =   3480
      TabIndex        =   40
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   19
      Left            =   1200
      TabIndex        =   19
      Text            =   "20"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   18
      Left            =   840
      TabIndex        =   18
      Text            =   "5"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   17
      Left            =   480
      TabIndex        =   17
      Text            =   "18"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   16
      Left            =   120
      TabIndex        =   16
      Text            =   "17"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   15
      Left            =   1200
      TabIndex        =   15
      Text            =   "4"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   14
      Left            =   840
      TabIndex        =   14
      Text            =   "19"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   13
      Left            =   480
      TabIndex        =   13
      Text            =   "6"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   12
      Left            =   120
      TabIndex        =   12
      Text            =   "16"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   11
      Left            =   1200
      TabIndex        =   11
      Text            =   "11"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   10
      Left            =   840
      TabIndex        =   10
      Text            =   "3"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   9
      Left            =   480
      TabIndex        =   9
      Text            =   "15"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Text            =   "7"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   7
      Left            =   1200
      TabIndex        =   7
      Text            =   "12"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   6
      Left            =   840
      TabIndex        =   6
      Text            =   "10"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Text            =   "2"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Text            =   "8"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Text            =   "13"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Text            =   "14"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Text            =   "9"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtEntrada 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   19
      Left            =   2880
      TabIndex        =   39
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   18
      Left            =   2520
      TabIndex        =   38
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   17
      Left            =   2160
      TabIndex        =   37
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   16
      Left            =   1800
      TabIndex        =   36
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   15
      Left            =   2880
      TabIndex        =   35
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   14
      Left            =   2520
      TabIndex        =   34
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   13
      Left            =   2160
      TabIndex        =   33
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   12
      Left            =   1800
      TabIndex        =   32
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   11
      Left            =   2880
      TabIndex        =   31
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   10
      Left            =   2520
      TabIndex        =   30
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   9
      Left            =   2160
      TabIndex        =   29
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   8
      Left            =   1800
      TabIndex        =   28
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   7
      Left            =   2880
      TabIndex        =   27
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   6
      Left            =   2520
      TabIndex        =   26
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   5
      Left            =   2160
      TabIndex        =   25
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   1800
      TabIndex        =   24
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   23
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   22
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   21
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   20
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub swap(bytx As Variant, byty As Variant)
Dim bytZ As Variant

bytZ = bytx
bytx = byty
byty = bytZ

End Sub

Private Sub ordenar(ByRef vector() As Byte)
Dim bytn As Byte
Dim byto As Byte

For bytn = 1 To 19
    For byto = 0 To 18
        If vector(bytn) < vector(byto) Then
            swap vector(bytn), vector(byto)
        End If
    Next byto
Next bytn
End Sub
Private Sub espiral(vector() As Byte, matriz() As Byte)
Dim byta As Byte
Dim bytn As Byte
Dim fila As Integer
Dim col As Integer
Dim contador As Byte

Do: DoEvents
    For col = bytn To 3 - bytn
        fila = bytn
        matriz(fila, col) = vector(contador)
        contador = contador + 1
    Next col
    For fila = 1 + bytn To 3 - bytn
        col = 3 - bytn
        matriz(fila, col) = vector(contador)
        contador = contador + 1
    Next fila
    For col = 3 - bytn To bytn Step -1
        fila = 4 - bytn
        matriz(fila, col) = vector(contador)
        contador = contador + 1
    Next col
    For fila = 3 - bytn To 1 + bytn Step -1
        col = bytn
        matriz(fila, col) = vector(contador)
        contador = contador + 1
    Next fila
bytn = bytn + 1
Loop Until (bytn > 1)
End Sub
Private Sub btnEspiral_Click()
Dim vector(19) As Byte
Dim matriz(4, 3) As Byte
Dim bytn As Byte
Dim fila As Byte
Dim col As Byte

For bytn = 0 To 19
    vector(bytn) = txtEntrada(bytn).Text
Next bytn
ordenar vector()
espiral vector(), matriz()
For bytn = 0 To 19
    fila = bytn \ 4
    col = bytn Mod 4
    lblSalida(bytn).Caption = matriz(fila, col)
Next bytn
End Sub
