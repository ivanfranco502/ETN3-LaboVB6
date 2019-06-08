VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnBinario 
      Caption         =   "Pasar a binario"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtBinario 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblCapri 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblBinario 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub caprichoso(ByRef vector() As Byte)
Dim bytn As Byte
Dim contador As Byte
For bytn = 0 To 7
    If vector(bytn) = 1 Then
        contador = contador + 1
    End If
Next bytn
If contador Mod 2 = 0 Then
    lblCapri.Caption = "No es caprichoso"
Else
    lblCapri.Caption = "Es caprichoso"
End If
End Sub
Private Sub Caprichoso1(ByRef vector1() As Byte)
Dim bytn As Byte
Dim contador As Byte
For bytn = 0 To 15
    If vector1(bytn) = 1 Then
        contador = contador + 1
    End If
Next bytn
If contador Mod 2 = 0 Then
    lblCapri.Caption = "No es caprichoso"
Else
    lblCapri.Caption = "Es caprichoso"
End If
End Sub
Private Sub btnBinario_Click()
Dim bytn As Integer
Dim valor As Integer



If IsNumeric(txtBinario.Text) Then
    valor = Val(txtBinario.Text)
    Select Case valor
        Case Is <= 128
            Dim vector(7) As Byte
            Dim resto(7) As Integer
            vector(0) = valor Mod 2
            resto(0) = valor \ 2
            bytn = 1
                
                Do: DoEvents
                    vector(bytn) = resto(bytn - 1) Mod 2
                    resto(bytn) = resto(bytn - 1) \ 2
                    bytn = bytn + 1
                Loop While (bytn <= 7)
                caprichoso vector()
                bytn = 7
                lblBinario.Caption = ""
                For bytn = 7 To 0 Step -1
                    lblBinario.Caption = lblBinario.Caption & vector(bytn)
                Next bytn
        Case Is <= 2568
        If valor > 128 Then
            Dim vector1(15) As Byte
            Dim resto1(15) As Byte
            vector1(0) = valor Mod 2
            resto1(0) = valor \ 2
            bytn = 1
                
                Do: DoEvents
                    vector1(bytn) = resto1(bytn - 1) Mod 2
                    resto1(bytn) = resto1(bytn - 1) \ 2
                    bytn = bytn + 1
                Loop While (bytn <= 7)
                caprichoso vector()
                bytn = 7
                lblBinario.Caption = ""
                For bytn = 7 To 0 Step -1
                    lblBinario.Caption = lblBinario.Caption & vector1(bytn)
                Next bytn
        End If
        Case Is > 256
            Dim vector2(31)
            Dim resto2(31)
            vector2(0) = valor Mod 2
            resto2(0) = valor \ 2
            bytn = 1
                
                Do: DoEvents
                    vector2(bytn) = resto2(bytn - 1) Mod 2
                    resto2(bytn) = resto2(bytn - 1) \ 2
                    bytn = bytn + 1
                Loop While (bytn <= 15)
                Caprichoso1 vector()
                bytn = 7
                lblBinario.Caption = ""
                For bytn = 15 To 0 Step -1
                    lblBinario.Caption = lblBinario.Caption & vector2(bytn)
                Next bytn

    End Select
End If
End Sub
