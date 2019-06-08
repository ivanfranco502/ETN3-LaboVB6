VERSION 5.00
Begin VB.Form Ordenar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenar"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   2415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnFrecuencia 
      Caption         =   "Ordenar"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "x"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Ordenados"
      Height          =   255
      Left            =   1200
      TabIndex        =   23
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Random"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblEdad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   9
      Left            =   1080
      TabIndex        =   21
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label lblEdad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   8
      Left            =   1080
      TabIndex        =   20
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label lblEdad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   7
      Left            =   1080
      TabIndex        =   19
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblEdad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   6
      Left            =   1080
      TabIndex        =   18
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label lblEdad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   5
      Left            =   1080
      TabIndex        =   17
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblEdad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   1080
      TabIndex        =   16
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblEdad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   15
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblEdad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   14
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblEdad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblEdad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblSalida 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   9
      Left            =   240
      TabIndex        =   11
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblSalida 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   8
      Left            =   240
      TabIndex        =   10
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label lblSalida 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   9
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label lblSalida 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   8
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label lblSalida 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label lblSalida 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lblSalida 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblSalida 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblSalida 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblSalida 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "Ordenar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub segundo(ByRef bytv() As Byte, bytNum() As Byte, ByRef bytvec() As Byte)
Dim a As Byte
Dim b As Byte
    For a = 1 To 9
        For b = 0 To 8
            If bytNum(a) = bytNum(b) Then
            If bytvec(a) < bytvec(b) Then
                swap bytvec(a), bytvec(b)
            End If
            End If
        Next b
    Next a

For a = 0 To 9
    lblEdad(a).Caption = bytvec(a)
Next a
End Sub
Private Sub ordenar(ByRef bytv() As Byte, ByRef bytNum() As Byte)
Dim bytn As Byte
Dim bytl As Byte
Dim bytvec(9) As Byte
For bytn = 0 To 9
    For bytl = 0 To 9
        If bytNum(bytn) > bytNum(bytl) Then
            
            swap bytNum(bytn), bytNum(bytl)
            swap bytv(bytn), bytv(bytl)
            
        End If
    Next bytl
Next bytn

For bytn = 0 To 9
    bytvec(bytn) = bytv(bytn)
Next bytn
segundo bytv(), bytNum(), bytvec()
End Sub
Private Sub frecuencia(ByRef bytv() As Byte)
Dim bytNum(9) As Byte
Dim bytk As Byte
Dim bytj As Byte

For bytk = 0 To 9
    For bytj = 0 To 9
        If bytv(bytk) = bytv(bytj) Then
            bytNum(bytk) = bytNum(bytk) + 1
        End If
    Next bytj
Next bytk
ordenar bytv(), bytNum()

End Sub
Private Sub swap(ByRef varX As Variant, ByRef varY As Variant)
Dim varZ As Variant
    varZ = varX
    varX = varY
    varY = varZ
End Sub

Private Sub btnFrecuencia_Click()
Dim bytv(9) As Byte
Dim bytn As Byte

For bytn = 0 To 9
    bytv(bytn) = Fix(9 * Rnd)
Next bytn

For bytn = 0 To 9
    lblSalida(bytn).Caption = bytv(bytn)
Next bytn

frecuencia bytv()

End Sub

Private Sub btnSalir_Click()
End
End Sub

Private Sub Form_Load()
Randomize Timer
End Sub
