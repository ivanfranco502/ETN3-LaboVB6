VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "x"
      Height          =   255
      Left            =   6480
      TabIndex        =   42
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton btnMatriz 
      Caption         =   "Ordenar Matriz"
      Height          =   735
      Left            =   4680
      TabIndex        =   41
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton btnFila 
      Caption         =   "Ordenar Filas"
      Height          =   735
      Left            =   4680
      TabIndex        =   40
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   19
      Left            =   3360
      TabIndex        =   19
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   18
      Left            =   2280
      TabIndex        =   18
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   17
      Left            =   1200
      TabIndex        =   17
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   16
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   15
      Left            =   3360
      TabIndex        =   15
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   14
      Left            =   2280
      TabIndex        =   14
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   13
      Left            =   1200
      TabIndex        =   13
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   12
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   11
      Left            =   3360
      TabIndex        =   11
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   10
      Left            =   2280
      TabIndex        =   10
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   9
      Left            =   1200
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   7
      Left            =   3360
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtMatriz 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   19
      Left            =   3360
      TabIndex        =   39
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   18
      Left            =   2280
      TabIndex        =   38
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   17
      Left            =   1200
      TabIndex        =   37
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   16
      Left            =   120
      TabIndex        =   36
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   15
      Left            =   3360
      TabIndex        =   35
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   14
      Left            =   2280
      TabIndex        =   34
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   13
      Left            =   1200
      TabIndex        =   33
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   12
      Left            =   120
      TabIndex        =   32
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   11
      Left            =   3360
      TabIndex        =   31
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   10
      Left            =   2280
      TabIndex        =   30
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   9
      Left            =   1200
      TabIndex        =   29
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   28
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   7
      Left            =   3360
      TabIndex        =   27
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   26
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   5
      Left            =   1200
      TabIndex        =   25
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   24
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   23
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   22
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   21
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblMatriz 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   975
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
Private Sub OrdenarCompleto(ByRef Matriz() As Byte)
Dim byta As Byte
Dim bytb As Byte
Dim fila1 As Byte
Dim fila2 As Byte
Dim col1 As Byte
Dim col2 As Byte

For byta = 0 To 18
    fila1 = byta \ 4
    col1 = byta Mod 4
    For bytb = (byta + 1) To 19
        fila2 = bytb \ 4
        col2 = bytb Mod 4
        If Matriz(fila1, col1) > Matriz(fila2, col2) Then
            swap Matriz(fila1, col1), Matriz(fila2, col2)
        End If
    Next bytb
Next byta

End Sub

Private Sub Ordenar(ByRef Matriz() As Byte)
Dim byta As Byte
Dim bytb As Byte
Dim fila1 As Byte
Dim fila2 As Byte
Dim col1 As Byte
Dim col2 As Byte

For byta = 0 To 19
    fila1 = byta \ 4
    col1 = byta Mod 4
    For bytb = 0 To 3
        If Matriz(fila1, col1) < Matriz(fila1, bytb) Then
            swap Matriz(fila1, col1), Matriz(fila1, bytb)
        End If
    Next bytb
Next byta

End Sub

Private Sub btnFila_Click()
Dim bytn As Byte
Dim matrix(4, 3) As Byte
Dim fila As Byte
Dim col As Byte
Dim pos As Byte

    For bytn = 0 To 19
        pos = bytn
        fila = pos \ 4
        col = pos Mod 4
        matrix(fila, col) = txtMatriz(bytn).Text
    Next bytn
Ordenar matrix()
    For bytn = 0 To 19
        pos = bytn
        fila = pos \ 4
        col = pos Mod 4
        lblMatriz(bytn).Caption = matrix(fila, col)
    Next bytn
End Sub

Private Sub btnMatriz_Click()
Dim bytn As Byte
Dim matrix(4, 3) As Byte
Dim fila As Byte
Dim col As Byte
Dim pos As Byte

For bytn = 0 To 19
pos = bytn
fila = pos \ 4
col = pos Mod 4
   matrix(fila, col) = txtMatriz(bytn).Text
Next bytn
OrdenarCompleto matrix()
For bytn = 0 To 19
pos = bytn
fila = pos \ 4
col = pos Mod 4
    lblMatriz(bytn).Caption = matrix(fila, col)
Next bytn
End Sub

Private Sub btnSalir_Click()
End
End Sub

