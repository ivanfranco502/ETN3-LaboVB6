VERSION 5.00
Begin VB.Form TP5 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Van's Rome Numbers"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnArabe 
      BackColor       =   &H00808080&
      Caption         =   "Transformar a Romano"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtArabe 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton btnSalir 
      BackColor       =   &H00808080&
      Cancel          =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton btnRomano 
      BackColor       =   &H00808080&
      Caption         =   "Transformar a Arabigo"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtRomano 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ingrese un numero  arabigo y se le mostrara el numero romano correspondiente."
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2400
      Y1              =   0
      Y2              =   2160
   End
   Begin VB.Label lblArabe 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ingrese un numero romano y se le mostrara el numero arabigo correspondiente"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "TP5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function intArabigo(ByVal strT As String) As Integer
    Select Case UCase(strT)
        Case Is = "I"
            intArabigo = 1
        Case Is = "V"
            intArabigo = 5
        Case Is = "X"
            intArabigo = 10
        Case Is = "L"
            intArabigo = 50
        Case Is = "C"
            intArabigo = 100
        Case Is = "D"
            intArabigo = 500
        Case Is = "M"
            intArabigo = 1000
        Case Else
            intArabigo = 0
    End Select
End Function


Private Sub btnArabe_Click()
Dim strArabe As String
Dim bytLetras As Byte
Dim strLetra As String
Dim bytN As Byte
Dim bytLenght As Byte
    
    txtArabe.Text = UCase(txtArabe.Text)
    strArabe = txtArabe.Text
    bytLetras = Len(strArabe)
    bytN = 0
    bytLenght = bytLetras + 1
    
        For bytN = 1 To bytLetras Step 1
            bytLenght = bytLenght - 1
            Select Case Mid(strArabe, bytN, 1)
                Case Is = "1"
                    Select Case bytLenght
                        Case Is = "4"
                            strLetra = strLetra & "M"
                        Case Is = "3"
                            strLetra = strLetra & "C"
                        Case Is = "2"
                            strLetra = strLetra & "X"
                        Case Is = "1"
                            strLetra = strLetra & "I"
                    End Select
                Case Is = "2"
                    Select Case bytLenght
                        Case Is = "4"
                                strLetra = strLetra & "MM"
                            Case Is = "3"
                                strLetra = strLetra & "CC"
                            Case Is = "2"
                                strLetra = strLetra & "XX"
                            Case Is = "1"
                                strLetra = strLetra & "II"
                    End Select
                Case Is = "3"
                    Select Case bytLenght
                        Case Is = "4"
                                strLetra = strLetra & "MMM"
                            Case Is = "3"
                                strLetra = strLetra & "CCC"
                            Case Is = "2"
                                strLetra = strLetra & "XXX"
                            Case Is = "1"
                                strLetra = strLetra & "III"
                     End Select
                Case Is = "4"
                        Select Case bytLenght
                            Case Is = "4"
                                strLetra = strLetra & "MMMM"
                            Case Is = "3"
                                strLetra = strLetra & "LC"
                            Case Is = "2"
                                strLetra = strLetra & "XL"
                            Case Is = "1"
                                strLetra = strLetra & "IV"
                        End Select
                Case Is = "5"
                    Select Case bytLenght
                        Case Is = "4"
                                strLetra = strLetra & "MMMMMM"
                            Case Is = "3"
                                strLetra = strLetra & "D"
                            Case Is = "2"
                                strLetra = strLetra & "L"
                            Case Is = "1"
                                strLetra = strLetra & "V"
                    End Select
                Case Is = "6"
                    Select Case bytLenght
                        Case Is = "4"
                                strLetra = strLetra & "MMMMMMM"
                            Case Is = "3"
                                strLetra = strLetra & "DC"
                            Case Is = "2"
                                strLetra = strLetra & "LX"
                            Case Is = "1"
                                strLetra = strLetra & "VI"
                    End Select
                Case Is = "7"
                    Select Case bytLenght
                        Case Is = "4"
                                strLetra = strLetra & "MMMMMMM"
                            Case Is = "3"
                                strLetra = strLetra & "DCC"
                            Case Is = "2"
                                strLetra = strLetra & "LXX"
                            Case Is = "1"
                                strLetra = strLetra & "VII"
                    End Select
                Case Is = "8"
                Select Case bytLenght
                        Case Is = "4"
                                strLetra = strLetra & "MMMMMMMM"
                            Case Is = "3"
                                strLetra = strLetra & "DCCC"
                            Case Is = "2"
                                strLetra = strLetra & "LXXX"
                            Case Is = "1"
                                strLetra = strLetra & "VIII"
                    End Select
                Case Is = "9"
                Select Case bytLenght
                        Case Is = "4"
                                strLetra = strLetra & "MMMMMMMMM"
                            Case Is = "3"
                                strLetra = strLetra & "CM"
                            Case Is = "2"
                                strLetra = strLetra & "XC"
                            Case Is = "1"
                                strLetra = strLetra & "IX"
                    End Select
                Case Is = "0"
            Case Else
                lblArabe.Caption = "Error de sintaxis"
            End Select
        Next bytN
        lblArabe.Caption = strLetra
End Sub

Private Sub btnRomano_Click()

Dim bytN As Byte
Dim intNumArab As Integer
Dim intx As Integer

    intNumArab = intArabigo(Right(txtRomano.Text, 1))
    bytN = 0
    
        For bytN = 1 To Len(txtRomano.Text) - 1 Step 1
        intx = intArabigo(Mid(txtRomano.Text, bytN, 1))
            If intx >= intArabigo(Mid(txtRomano.Text, bytN + 1, 1)) Then
                intNumArab = intNumArab + intArabigo(Mid(txtRomano.Text, bytN, 1))
            Else
                intNumArab = intNumArab - intArabigo(Mid(txtRomano.Text, bytN, 1))
            End If
        Next bytN
            lblArabe.Caption = intNumArab
End Sub

Private Sub btnSalir_Click()
End
End Sub

