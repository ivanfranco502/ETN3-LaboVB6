VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lista 
      Height          =   3765
      ItemData        =   "Tropa.frx":0000
      Left            =   2880
      List            =   "Tropa.frx":0002
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtBuscar 
      Height          =   405
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtPalabra 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton btnbuscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton salir 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton btnTropa 
      Caption         =   "Ver Probabilidades"
      Default         =   -1  'True
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Posicion de la palabra en busqueda"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Busqueda de palabra"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Inserte una palabra de >= 5"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vecTrapoReal(199) As String
Private Sub crear(ByRef vector() As String, v() As String)
Dim byta As Byte
Dim bytb As Byte
Dim palabra As String
    For byta = 0 To 199
    palabra = ""
        For bytb = 1 To 5
            Select Case (Mid(vector(byta), bytb, 1))
                Case Is = 0
                    palabra = palabra & v(0)
                Case Is = 1
                    palabra = palabra & v(1)
                Case Is = 2
                    palabra = palabra & v(2)
                Case Is = 3
                    palabra = palabra & v(3)
                Case Is = 4
                    palabra = palabra & v(4)
            End Select
        Next bytb
        vector(byta) = palabra
    Next byta
End Sub
Private Function boolEsta(ByVal palabra As String, ByRef vecTrapoReal() As String, ByVal bytt As Byte) As Boolean
Dim byta As Byte
Dim bytn As Byte
    For byta = 0 To bytt
    DoEvents
        If vecTrapoReal(byta) = palabra Then
            boolEsta = True
            Exit For
        Else
            boolEsta = False
        End If
    Next byta
End Function

Private Function boolrep(ByVal letra As Byte, ByVal palabra As String, ByVal i As Byte) As Boolean
Dim bytn As Byte

    For bytn = 1 To i
    DoEvents
        If Mid(palabra, bytn, 1) = letra Then
            boolrep = True
            Exit For
        Else
            boolrep = False
        End If
    Next bytn
End Function

Private Sub btnbuscar_Click()
Dim bytn As Byte

For bytn = 0 To 120
    If lista.List(bytn) = txtBuscar.Text Then
    lista.ListIndex = bytn
    Label1.Caption = bytn
    Exit For
    End If
Next bytn
btnbuscar.Enabled = False
End Sub

Private Sub btnTropa_Click()
Dim byta As Byte
Dim i As Byte
Dim palabra As String
Dim bytn As Byte
Dim letra As Byte
Dim bytt As Long
Dim bytcontador As Byte
Dim intA As Integer
Dim alert As VbMsgBoxResult
Dim vector(4) As String
Dim bytD As Byte
programa:
lista.Clear
Label1.Caption = ""
txtBuscar.Text = ""
btnbuscar.Enabled = False
If txtPalabra.Text = "" Then
alert = MsgBox("Ingrese una palabra", vbOKOnly, "Sin palabra")
    If alert = vbOK Then
        GoTo programa
    End If
Else
bytD = 1
For bytt = 0 To 4
    vector(bytt) = Mid(txtPalabra.Text, bytD, 1)
    bytD = bytD + 1
Next bytt
For bytt = 0 To 10000
DoEvents
palabra = ""
For i = 1 To 5
    Do: DoEvents
        letra = Int(Rnd * 5)
    Loop While (boolrep(letra, palabra, i) = True)
    
    palabra = palabra & letra
Next i
    If boolEsta(palabra, vecTrapoReal, bytcontador) = False Then
    bytcontador = bytcontador + 1
    vecTrapoReal(bytcontador) = palabra
    End If
Next bytt

crear vecTrapoReal(), vector()

For intA = 0 To 120
DoEvents
    lista.AddItem vecTrapoReal(intA)
Next intA
btnbuscar.Enabled = True
End If
End Sub

Private Sub Form_Load()
Randomize Timer
btnbuscar.Enabled = False
End Sub

Private Sub salir_Click()
End
End Sub
