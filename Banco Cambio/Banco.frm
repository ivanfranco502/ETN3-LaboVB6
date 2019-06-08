VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton btnCalcular 
      Caption         =   "Calcular Cambio"
      Default         =   -1  'True
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtMonto 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label mostrar 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "$ 50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   17
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label mostrar 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "$ 100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label mostrar 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "$ 20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label mostrar 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "$ 10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label mostrar 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "$ 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4920
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label mostrar 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "$ 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6120
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label mostrar 
      Caption         =   "Billetes"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label mostrar 
      Caption         =   "Cant Veces"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label resultado 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label resultado 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label resultado 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   2640
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label resultado 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   3840
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label resultado 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   5040
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label resultado 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   5
      Left            =   6240
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Ingrese un monto par:"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Ver()
Dim bytn As Byte
    For bytn = 0 To 7
        mostrar(bytn).Visible = True
    Next bytn
    For bytn = 0 To 5
        resultado(bytn).Visible = True
    Next bytn
End Sub
Private Sub Ocultar()
Dim bytn As Byte
    For bytn = 0 To 7
        mostrar(bytn).Visible = False
    Next bytn
    For bytn = 0 To 5
    resultado(bytn).Visible = False
    Next bytn
End Sub
Private Sub btnCalcular_Click()
Dim alert As VbMsgBoxResult
Ocultar
If IsNumeric(txtMonto.Text) Then
   
        Dim cien As Byte
        Dim cincuenta As Byte
        Dim veinte As Byte
        Dim diez As Byte
        Dim cinco As Byte
        Dim dos As Byte
        Dim resto(1) As Integer
        Dim intValor As Integer
        Dim sumatotal As Integer
        
        intValor = Val(txtMonto.Text)
        
        cien = Fix(intValor / 2 / 100)
            resto(0) = intValor - (100 * cien)
        cincuenta = Fix(resto(0) / 2 / 50)
            resto(1) = resto(0) - (50 * cincuenta)
        veinte = Fix(resto(1) / 2 / 20)
            resto(0) = resto(1) - (20 * veinte)
        diez = Fix(resto(0) / 2 / 10)
            resto(1) = resto(0) - (10 * diez)
        cinco = Fix(resto(1) / 2 / 5)
            resto(0) = resto(1) - (5 * cinco)
        If resto(0) Mod 2 = 0 Then
            dos = Fix(resto(0) / 2)
        Else
            resto(0) = resto(0) - 5
                If resto(0) < 0 Then
                    alert = MsgBox("Valor Imposible de dar cambio en billetes, Ingrese un nuevo valor", vbOKOnly, "Valor Erroneo")
                        If alert = vbOK Then
                            txtMonto.Text = ""
                            GoTo fin
                        End If
                End If
            cinco = cinco + 1
            dos = Fix(resto(0) / 2)
        End If
        Ver
        resultado(0) = cien
        resultado(1) = cincuenta
        resultado(2) = veinte
        resultado(3) = diez
        resultado(4) = cinco
        resultado(5) = dos

Else
alert = MsgBox("El monto ingresado no es numérico, intentelo de nuevo", vbOKOnly, "Monto Erróneo")
    If alert = vbOK Then
        txtMonto.Text = ""
    End If
End If
fin:
End Sub


Private Sub btnSalir_Click()
End
End Sub

