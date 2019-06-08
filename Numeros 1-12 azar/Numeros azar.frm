VERSION 5.00
Begin VB.Form NumAzar 
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "X"
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton btnGenerar 
      Caption         =   "Generar Numeros al azar  1 - 12"
      Height          =   735
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblNumero 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   9
      Left            =   4920
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblNumero 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   8
      Left            =   3720
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblNumero 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   7
      Left            =   2520
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblNumero 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblNumero 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblNumero 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   4920
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblNumero 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblNumero 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblNumero 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblNumero 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "NumAzar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function boolNum(ByRef bytVector() As Byte, ByVal bytX As Byte, bytA As Byte)
Dim bytY As Byte
Dim boolTrue As Boolean

bytY = 0

Do: DoEvents


If bytA = bytVector(bytY) Then
        boolTrue = True
End If
bytY = bytY + 1
Loop Until (bytY = bytX Or boolTrue = True)
boolNum = boolTrue
End Function
Private Sub btnGenerar_Click()
Dim bytV(9) As Byte
Dim bytN As Byte
Dim bytAux As Byte

bytV(0) = Fix((12) * Rnd) + 1
lblNumero(0).Caption = bytV(0)
For bytN = 1 To 9
    Do: DoEvents
        bytAux = Fix((12) * Rnd) + 1
        
    Loop While (boolNum(bytV(), bytN, bytAux))
    bytV(bytN) = bytAux
    lblNumero(bytN).Caption = bytV(bytN)
Next bytN

End Sub

Private Sub btnSalir_Click()
End
End Sub

Private Sub Form_Load()
Randomize Timer
End Sub
