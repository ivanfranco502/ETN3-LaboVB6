VERSION 5.00
Begin VB.Form TP1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Van's Date Calculator v.1.0"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDia 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtMes 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton btnSalir 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton btnCalcular 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Calcular"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtAnio 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dia"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblAño 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.Label LblMes 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mes"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblSalida 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   2895
   End
End
Attribute VB_Name = "TP1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCalcular_Click()
Dim intAnio As Integer, bytMes As Byte, bytCantDias As Byte, bytDia As Byte
Dim bytA As Byte, intB As Integer, bytC As Byte, bytD As Byte, bytxE As Byte
Dim intF As Integer, lonxG As Long, intH As Integer, lonI As Long, bytJ As Byte
    intAnio = Val(txtAnio.Text)     'Se consigue el dato del año
    bytMes = Val(txtMes.Text)       'Se consigue el dato del mes
    bytDia = Val(txtDia.Text)       'Se consigue el dato del dia
    lblSalida.Caption = ""          'Se borra el Label de salida
If IsNumeric(txtAnio.Text) Then     'Se confirma el valor del año
    Select Case (bytMes)
        Case Is = 1, 3, 5, 7, 8, 10, 12 'Se calcula los meses con 31 dias
            bytCantDias = 31
        Case Is = 4, 6, 9, 11           'Se calcula los meses con 30 dias
            bytCantDias = 30
        Case Is = 2                     'Se calcula el año biciesto
            If (intAnio Mod 100 = 0) Then
                    If (intAnio Mod 400 = 0) Then
                        bytCantDias = 29
                    Else
                        bytCantDias = 28
                    End If
            Else
                    If (intAnio Mod 4 = 0) Then
                        bytCantDias = 29
                    Else
                        bytCantDias = 28
                    End If
            End If
        Case Else
            lblSalida.Caption = "Valor incorrecto: Mes" 'Error en Mes
    End Select
'Algoritmo de Pascal, para calcular los dias
        
        If lblSalida.Caption = "" Then
            bytA = Fix((12 - bytMes) / 10)
            intB = intAnio - bytA
            bytC = bytMes + 12 * bytA
            bytD = Fix(intB / 100)
            bytxE = Fix(bytD / 4)
            intF = 2 - bytD + bytxE
            lonxG = Fix(365.25 * intB)
            intH = Fix(30.6001 * (bytC + 1))
            lonI = intF + lonxG + intH + bytDia + 5
            bytJ = lonI Mod 7
                If (bytDia >= 1 And bytDia <= bytCantDias) Then
                        Select Case bytJ 'Calculo del dia
                            Case Is = 0
                                lblSalida.Caption = "Sábado"
                            Case Is = 1
                                lblSalida.Caption = "Domingo"
                            Case Is = 2
                                lblSalida.Caption = "Lunes"
                            Case Is = 3
                                lblSalida.Caption = "Martes"
                            Case Is = 4
                                lblSalida.Caption = "Miercoles"
                            Case Is = 5
                                lblSalida.Caption = "Jueves"
                            Case Is = 6
                                lblSalida.Caption = "Viernes"
                        End Select
                Else
                    lblSalida.Caption = "Valor Incorrecto: Dia" 'Error Dia
                End If
        End If
Else
    lblSalida.Caption = "Valor Incorrecto: Año" 'Error Año
End If

End Sub


Private Sub btnSalir_Click()
End
End Sub

