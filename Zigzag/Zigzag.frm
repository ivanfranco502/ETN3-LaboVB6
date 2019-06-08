VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnZigzag 
      Caption         =   "Ordenar en Zigzag"
      Height          =   495
      Left            =   720
      TabIndex        =   18
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtZig 
      Height          =   375
      Index           =   8
      Left            =   1080
      TabIndex        =   8
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtZig 
      Height          =   375
      Index           =   7
      Left            =   600
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtZig 
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtZig 
      Height          =   375
      Index           =   5
      Left            =   1080
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtZig 
      Height          =   375
      Index           =   4
      Left            =   600
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtZig 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtZig 
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtZig 
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtZig 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblZig 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   8
      Left            =   2760
      TabIndex        =   17
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblZig 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   7
      Left            =   2280
      TabIndex        =   16
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblZig 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   6
      Left            =   1800
      TabIndex        =   15
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblZig 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   5
      Left            =   2760
      TabIndex        =   14
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblZig 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   2280
      TabIndex        =   13
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblZig 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   12
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblZig 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblZig 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblZig 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   495
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
Private Sub filas(ByRef matriz() As Byte)
Dim byta As Byte
Dim bytb As Byte
Dim fila1 As Byte
Dim col1 As Byte


For byta = 3 To 5
    fila1 = byta \ 3
    col1 = byta Mod 3
    For bytb = 0 To 2
        If matriz(fila1, col1) > matriz(fila1, bytb) Then
            swap matriz(fila1, col1), matriz(fila1, bytb)
        End If
    Next bytb
Next byta

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
filas matriz()
End Sub

Private Sub btnZigzag_Click()
Dim bytn As Byte
Dim pos As Byte
Dim Fila As Byte
Dim col As Byte
Dim matriz(2, 2) As Byte

For bytn = 0 To 8
    pos = bytn
    Fila = pos \ 3
    col = pos Mod 3
    matriz(Fila, col) = txtZig(bytn).Text
Next bytn
ordenar matriz()
For bytn = 0 To 8
    pos = bytn
    Fila = pos \ 3
    col = pos Mod 3
    lblZig(bytn).Caption = matriz(Fila, col)
Next bytn
End Sub

