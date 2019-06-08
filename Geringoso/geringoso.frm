VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnReal 
      Caption         =   "Pasar a Real"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton BtnGeringoso 
      Caption         =   "Pasar a Geringoso"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtPalabra 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblGeringoso 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Geringoso:"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Inserte una palabra"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function real(ByRef palabra As String) As String
Dim bytn As Byte
Dim strReal As String
Dim bytvalor As Byte
bytvalor = Len(palabra)
    For bytn = 1 To bytvalor
        Select Case (Mid(palabra, bytn, 1))
            Case Is = "a"
                If (Mid(palabra, bytn + 1, 2)) = "pa" Then
                    palabra = Left(palabra, bytn) & Right(palabra, Len(palabra) - (Len(Left(palabra, bytn)) + 2))
                End If
            Case Is = "e"
                If (Mid(palabra, bytn + 1, 2)) = "pe" Then
                    palabra = Left(palabra, bytn) & Right(palabra, Len(palabra) - (Len(Left(palabra, bytn)) + 2))
                End If
            Case Is = "i"
                If (Mid(palabra, bytn + 1, 2)) = "pi" Then
                    palabra = Left(palabra, bytn) & Right(palabra, Len(palabra) - (Len(Left(palabra, bytn)) + 2))
                End If
            Case Is = "o"
                If (Mid(palabra, bytn + 1, 2)) = "po" Then
                    palabra = Left(palabra, bytn) & Right(palabra, Len(palabra) - (Len(Left(palabra, bytn)) + 2))
                End If
            Case Is = "u"
                If (Mid(palabra, bytn + 1, 2)) = "pu" Then
                    palabra = Left(palabra, bytn) & Right(palabra, Len(palabra) - (Len(Left(palabra, bytn)) + 2))
                End If
        End Select
    Next bytn
strReal = palabra
real = strReal
                
End Function
Private Function separador(ByRef palabra As String) As String
Dim bytn As Byte
Dim Geringoso As String
Dim bytvalor As Byte
bytvalor = Len(palabra)
    For bytn = 1 To bytvalor
        Select Case (Mid(palabra, bytn, 1))
            Case Is = "a"
                Geringoso = Geringoso & Mid(palabra, bytn, 1) & "pa"
            Case Is = "e"
                Geringoso = Geringoso & Mid(palabra, bytn, 1) & "pe"
            Case Is = "i"
               Geringoso = Geringoso & Mid(palabra, bytn, 1) & "pi"
            Case Is = "o"
                Geringoso = Geringoso & Mid(palabra, bytn, 1) & "po"
            Case Is = "u"
                Geringoso = Geringoso & Mid(palabra, bytn, 1) & "pu"
            Case Else
                Geringoso = Geringoso & Mid(palabra, bytn, 1)
        End Select
    Next bytn
separador = Geringoso
End Function
Private Sub BtnGeringoso_Click()
Dim palabra As String
palabra = LCase(txtPalabra.Text)

lblGeringoso.Caption = separador(palabra)

End Sub

Private Sub btnReal_Click()
Dim palabra As String
palabra = LCase(txtPalabra.Text)
lblGeringoso.Caption = real(palabra)

End Sub
