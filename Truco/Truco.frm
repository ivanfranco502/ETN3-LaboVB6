VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnEnvido 
      Caption         =   "Envido"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblCartas 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Maso
    strPalo As String * 8
    bytvalor As Byte
End Type
Private Sub swap(ByRef a As Variant, ByRef b As Variant)
Dim c As Variant

c = a
a = b
b = c

End Sub
Private Sub Envido(ByRef mazo() As Maso)
Dim bytn As Byte
Dim Palo(3) As Byte
Dim bytEnvido As Byte
Dim bytM As Byte
Dim strPalo As String * 8
Dim boolTres As Boolean

For bytn = 0 To 2
    Select Case mazo(bytn).strPalo
        Case Is = "Oro     "
            Palo(0) = Palo(0) + 1
        Case Is = "Espada  "
            Palo(1) = Palo(1) + 1
        Case Is = "Copa    "
            Palo(2) = Palo(2) + 1
        Case Is = "Basto   "
            Palo(3) = Palo(3) + 1
    End Select
    Print mazo(bytn).bytvalor & " - " & mazo(bytn).strPalo
Next bytn
bytn = 0
Do: DoEvents
    If Palo(bytn) > 1 Then
        bytEnvido = 20
    End If
bytn = bytn + 1
Loop Until (bytEnvido = 20 Or bytn > 3)

If bytEnvido = 20 Then
    For bytM = 0 To 3
        If Palo(bytM) = 2 Then
                Select Case bytM
                    Case Is = 0
                        strPalo = "Oro     "
                    Case Is = 1
                        strPalo = "Espada  "
                    Case Is = 2
                        strPalo = "Copa    "
                    Case Is = 3
                        strPalo = "Basto   "
                End Select
        ElseIf Palo(bytM) = 3 Then
            boolTres = True
            Select Case bytM
                Case Is = 0
                    strPalo = "Oro     "
                Case Is = 1
                    strPalo = "Espada  "
                Case Is = 2
                    strPalo = "Copa    "
                Case Is = 3
                    strPalo = "Basto   "
            End Select
        End If
    Next bytM
        If boolTres = False Then
            For bytn = 0 To 2
                If mazo(bytn).bytvalor <= 7 Then
                    If mazo(bytn).strPalo = strPalo Then
                        bytEnvido = bytEnvido + mazo(bytn).bytvalor
                    End If
                End If
                lblCartas.Caption = "Envido" & "  " & bytEnvido
            Next bytn
        Else
            For bytn = 0 To 2
                If mazo(bytn).bytvalor <= 7 Then
                    If mazo(bytn).strPalo = strPalo Then
                        bytEnvido = bytEnvido + mazo(bytn).bytvalor
                    End If
                End If
                lblCartas.Caption = "Flor" & "  " & bytEnvido
            Next bytn
        End If
Else
    lblCartas.Caption = "No Hay Envido"
End If

End Sub
Private Sub DesordenarMazo(ByRef mazo() As Maso)
Dim bytn As Byte
Dim azar As Byte

For bytn = 0 To 39
    azar = Fix(Rnd * 40)
    swap mazo(bytn).bytvalor, mazo(azar).bytvalor
    swap mazo(bytn).strPalo, mazo(azar).strPalo
Next bytn
End Sub

Private Sub btnEnvido_Click()
Dim masoCarta(39) As Maso
Dim bytn As Byte
Dim bytC As Byte
Dim bytD As Byte
bytC = 1

Form1.Cls
lblCartas.Caption = ""
For bytn = 0 To 39 Step 10
    For bytD = bytn To bytn + 9
        If bytC = 8 Then
            bytC = bytC + 2
        End If
        masoCarta(bytD).bytvalor = bytC
        bytC = bytC + 1
    Next bytD

bytC = 1
Next bytn

For bytn = 0 To 39
    If bytn >= 0 And bytn <= 9 Then
        masoCarta(bytn).strPalo = "Oro"
    ElseIf bytn >= 10 And bytn <= 19 Then
        masoCarta(bytn).strPalo = "Espada"
    ElseIf bytn >= 20 And bytn <= 29 Then
        masoCarta(bytn).strPalo = "Basto"
    Else
        masoCarta(bytn).strPalo = "Copa"
    End If
Next bytn

DesordenarMazo masoCarta()
Envido masoCarta()
End Sub

Private Sub Form_Load()
Randomize Timer
End Sub
