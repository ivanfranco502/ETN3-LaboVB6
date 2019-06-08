VERSION 5.00
Begin VB.Form TP3 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Van's Word Inverse"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCapicua 
      BackColor       =   &H00808080&
      Caption         =   "&Capicua"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton btnInvertir 
      BackColor       =   &H00808080&
      Caption         =   "&Invertir"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtPalabra 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label lblInvertida 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ingrese una palabra y esta sera dada vuelta y mostrada."
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "TP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function boolEsPalin(ByVal strT As String) As Boolean
    
    If strT = strInvertir(strT) Then
        boolEsPalin = True
    Else
        boolEsPalin = False
    End If
        
End Function
Private Function strHacerPalin(ByVal strT As String) As String
Dim strX As String, bytCantCaract As Byte
    
    bytCantCaract = 0
    Do: DoEvents
        strX = "": bytCantCaract = bytCantCaract + 1
        strX = strT & strInvertir(Left(strT, bytCantCaract))
    Loop While Not (boolEsPalin(strX))
    strHacerPalin = strX
    
End Function
Private Function strInvertir(ByVal strT As String) As String
Dim strX As String: strX = ""
Dim intPosi As Integer
    
    For intPosi = Len(strT) To 1 Step -1
        strX = strX & Mid(strT, intPosi, 1)
    Next intPosi
    strInvertir = strX
        
End Function

Private Sub btnCapicua_Click()

    If (boolEsPalin(txtPalabra.Text)) Then
        lblInvertida.Caption = "Es palindromo la palabra ingresada"
    Else
        lblInvertida.Caption = strHacerPalin(txtPalabra.Text)
    End If

End Sub

Private Sub btnInvertir_Click()
Dim strPalabra As String, bytLetras As Byte, intRestador As Integer

bytLetras = Len(txtPalabra.Text)
lblInvertida.Caption = ""
    For intRestador = bytLetras To 1 Step -1
        strPalabra = strPalabra & Mid(txtPalabra.Text, intRestador, 1)
    Next intRestador
        lblInvertida.Caption = strPalabra
            
End Sub
