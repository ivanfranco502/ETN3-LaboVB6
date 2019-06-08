VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnVectores 
      Caption         =   "Vectores al azar"
      Height          =   615
      Left            =   4560
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton btnLlenar 
      Caption         =   "Llenar"
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton btnEnd 
      Cancel          =   -1  'True
      Caption         =   "x"
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton btnCaracteres 
      Caption         =   "Realizar azar"
      Default         =   -1  'True
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtCarac 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblAzar 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   6105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function Cadena(ByRef bytX As Byte) As String
Dim bytLet As Byte
Dim strPalabra As String
Dim bytCont As Byte
    Do: DoEvents
    
        bytLet = Fix((122 - 98) * Rnd) + 97
        strPalabra = strPalabra & Chr(bytLet)
        bytCont = bytCont + 1
    Loop While bytCont < bytX
    Cadena = strPalabra
lblAzar.Caption = Cadena
End Function
Private Sub btnLlenar_Click()
Static strPalabra(9) As String
Static bytPosi As Byte
    strPalabra(bytPosi) = Cadena(bytPosi)
    bytPosi = bytPosi + 1
    If bytPosi = 10 Then btnLlenar.Enabled = False
End Sub
Private Sub btnCaracteres_Click()
Dim bytCantCar
    bytCantCar = Val(txtCarac.Text)
    Cadena (bytCantCar)
End Sub
Private Sub btnEnd_Click()
End
End Sub
Private Sub Form_Load()
Randomize Timer
End Sub

