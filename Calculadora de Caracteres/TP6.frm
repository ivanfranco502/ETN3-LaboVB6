VERSION 5.00
Begin VB.Form TP6 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "x "
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton btnCalcular 
      Caption         =   "Calcular los porcentajes"
      Default         =   -1  'True
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtAlfa 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblPorcNum 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Porcentaje Numeros:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblPorcEsp 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Porcentaje Espacios:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblPorcCon 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Porcentaje Consonantes:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblPorcVoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Porcenaje Vocales:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
End
Attribute VB_Name = "TP6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCalcular_Click()
Dim str
strpalabra = txtAlfa.Text
bytCant = Len(strpalabra)
    For bytN = 1 To bytCant
        Select Case (Mid(strpalabra, bytCant, 1))
            Case Is = "b", "c", "d", "f", "g", "h", "j", "k", "l", "m", "n", "ñ" _
                      , "p", "q", "r", "s", "t", "v", "w", "x", "y", "z"
            
            Case Is = "a", "e", "i", "o", "u"
        
            Case Is = "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            
            Case Is = " "

End Sub

Private Sub btnSalir_Click()
End
End Sub
