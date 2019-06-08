VERSION 5.00
Begin VB.Form Patentes 
   Caption         =   "Generador de patentes al azar"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnPatentes 
      Caption         =   "Patentes"
      Default         =   -1  'True
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblPatente 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Patente"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
End
Attribute VB_Name = "Patentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnPatentes_Click()
Dim bytLet As Byte
Dim strPatente As String
Dim bytCont As Byte
Dim bytNum As Byte
Dim bytContNum As Byte


bytContNum = 1
bytCont = 1

Do: DoEvents
    
    bytLet = Fix((122 - 98) * Rnd) + 97
    strPatente = strPatente & Chr(bytLet)
    bytCont = bytCont + 1
Loop While bytCont <= 3
strPatente = strPatente & " "
Do: DoEvents
    bytNum = Fix((9) * Rnd)
    strPatente = strPatente & bytNum
    bytContNum = bytContNum + 1
Loop While bytContNum <= 3

lblPatente.Caption = UCase(strPatente)


End Sub

Private Sub Form_Load()
Randomize Timer
End Sub
