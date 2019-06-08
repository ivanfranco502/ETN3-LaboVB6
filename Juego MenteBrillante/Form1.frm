VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mente Brillante"
   ClientHeight    =   7230
   ClientLeft      =   4920
   ClientTop       =   2175
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Height          =   135
      Left            =   4320
      TabIndex        =   13
      Top             =   0
      Width           =   135
   End
   Begin VB.CommandButton btnGenerar 
      Caption         =   "Generar"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtNumero 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton btnJugar 
      Caption         =   "Jugar"
      Default         =   -1  'True
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblMalas 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label lblReg 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblBuenas 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3120
      TabIndex        =   17
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Malas:"
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Regulares:"
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Buenas:"
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblIngreso 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label lblIngreso 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label lblIngreso 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label lblIngreso 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblIngreso 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblIngreso 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Estado del ingreso:"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Numero ingresado:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblIngreso 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3720
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblAzar 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ventrada(4) As Byte
Dim Vcompu(4) As Byte
Dim Vres(4) As Byte

Private Function boolIgualdad(ByVal Number As Byte, ByVal Cadena As String) As Boolean
Dim bytN As Byte
Dim byta As Byte
Dim boolS As Boolean

For bytN = 1 To 5
    If (Mid(Cadena, bytN, 1)) = Number Then
        boolS = True
        Exit For
    End If
Next bytN

boolIgualdad = boolS
End Function
Private Sub Validar(ByVal Pal As String, Azar As String)
Dim bytN As Byte
Dim bytContBuenos As Byte
Dim bytContRegulares As Byte
Dim bytContMalos As Byte
Dim bytT As Byte

For bytN = 1 To 5
    If Mid(Pal, bytN, 1) = Mid(Azar, bytN, 1) Then
        bytContBuenos = bytContBuenos + 1
    Else
        For bytT = 1 To 5
            If Mid(Pal, bytN, 1) = Mid(Azar, bytT, 1) Then
                bytContRegulares = bytContRegulares + 1
            End If
        Next bytT
    End If
Next bytN
If (bytContBuenos + bytContRegulares) = 5 Then
    lblMalas.Caption = 0
Else
lblMalas.Caption = 5 - (bytContBuenos + bytContRegulares)
End If
lblBuenas.Caption = bytContBuenos
lblReg.Caption = bytContRegulares


End Sub
Private Sub Generar(ByRef num As String)

Dim bytN As Byte
Dim numer As Byte
Dim byta As Byte

    Do: DoEvents
        Do: DoEvents
            numer = Fix(8 * Rnd)
        Loop While (boolIgualdad(numer, num) = True)
        num = num & numer
        bytN = bytN + 1
    Loop While (bytN < 4)

        
End Sub
Private Sub btnGenerar_Click()
Dim numero As String
btnGenerar.Enabled = False
btnJugar.Enabled = True

numero = Fix(6 * Rnd + 1)
Generar numero

lblAzar.Caption = numero

End Sub

Private Sub btnJugar_Click()
Static bytN As Integer
Dim respuesta As VbMsgBoxResult
Static bytT As Integer
Dim strPal As String
Dim Azar As String
Dim ganaste As VbMsgBoxResult
Dim byta As Byte

If bytT < 7 Then
lblIngreso(bytT).Caption = txtNumero.Text
End If


Azar = lblAzar.Caption
strPal = txtNumero.Text

Validar strPal, Azar

If bytT >= 0 Then
txtNumero.Text = ""
End If

If lblBuenas.Caption = "5" Then
lblAzar.Visible = True
ganaste = MsgBox("GANASTE!!!DESEAS JUGAR DE NUEVO?", vbYesNo, "Juego Terminado")
        If ganaste = vbYes Then
            btnGenerar.Enabled = True
            btnJugar.Enabled = False
            bytN = -1
            bytT = -1
            For byta = 0 To 6
            lblIngreso(byta).Caption = ""
            Next byta
        Else
            btnSalir_Click
        End If
End If

bytT = bytT + 1
bytN = bytN + 1
If bytN = 7 Then
lblAzar.Visible = True
    respuesta = MsgBox("El juego ha terminado y no pudiste adivinar el numero que era " & Azar & ". DESEAS JUGAR DE NUEVO?", vbYesNo, "Juego Terminado")
        If respuesta = vbYes Then
            btnGenerar.Enabled = True
            btnJugar.Enabled = False
            bytN = -1
            bytT = -1
            For byta = 0 To 6
            lblIngreso(byta).Caption = ""
            Next byta
        Else
            btnSalir_Click
        End If
End If

End Sub

Private Sub btnSalir_Click()
End
End Sub

Private Sub Form_Load()
Randomize Timer
'con las letras de la palabra trapo
'formo todas las palabras tengan o no sentido
'de letras distintas
'sin letras repetidas
'en que posicion alfabetica se encuentra tropa?
btnJugar.Enabled = False
End Sub
Private Sub Salir_Click()
btnSalir_Click
End Sub
